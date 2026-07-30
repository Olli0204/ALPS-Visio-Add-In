using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    /// <summary>
    /// Trains a deterministic text classifier from the dataset bundled with
    /// the add-in. It deliberately avoids ML runtime dependencies so retraining
    /// remains reliable inside Visio's .NET Framework VSTO AppDomain.
    /// </summary>
    internal sealed class NlpNameClassifier
    {
        private const string TrainingResourceName =
            "ALPS_Visio_AddIn_rewrite.NlpChecking.training.tsv";

        private readonly object synchronization = new object();
        private IDictionary<string, BinaryClassModel> modelsByShapeType;
        private BinaryClassModel fallbackModel;
        private int trainingExampleCount;

        public int EnsureTrained()
        {
            lock (synchronization)
            {
                if (modelsByShapeType == null)
                    TrainCore();

                return trainingExampleCount;
            }
        }

        public int Retrain()
        {
            lock (synchronization)
            {
                TrainCore();
                return trainingExampleCount;
            }
        }

        public NlpNamePrediction Predict(string label, string shapeType)
        {
            lock (synchronization)
            {
                if (modelsByShapeType == null)
                    TrainCore();

                return PredictCore(label, shapeType);
            }
        }

        private void TrainCore()
        {
            IList<NlpTrainingExample> examples = ReadTrainingExamples();
            Dictionary<string, BinaryClassModel> newModels =
                new Dictionary<string, BinaryClassModel>(
                    StringComparer.OrdinalIgnoreCase);
            BinaryClassModel newFallbackModel = new BinaryClassModel();

            foreach (NlpTrainingExample example in examples)
            {
                if (!newModels.TryGetValue(
                    example.ShapeType, out BinaryClassModel model))
                {
                    model = new BinaryClassModel();
                    newModels.Add(example.ShapeType, model);
                }

                IList<string> features = ExtractFeatures(
                    example.Name, example.ShapeType).ToList();
                model.Add(example.IsValidName, features);
                newFallbackModel.Add(example.IsValidName, features);
            }

            foreach (BinaryClassModel model in newModels.Values)
                model.Validate();
            newFallbackModel.Validate();

            modelsByShapeType = newModels;
            fallbackModel = newFallbackModel;
            trainingExampleCount = examples.Count;
        }

        private NlpNamePrediction PredictCore(
            string label, string shapeType)
        {
            if (!modelsByShapeType.TryGetValue(
                shapeType ?? string.Empty, out BinaryClassModel model))
            {
                model = fallbackModel;
            }

            IList<string> features =
                ExtractFeatures(label, shapeType).ToList();
            int totalExamples = model.Valid.ExampleCount
                + model.Invalid.ExampleCount;

            double validScore = CalculateScore(
                model.Valid, model.VocabularySize,
                totalExamples, features);
            double invalidScore = CalculateScore(
                model.Invalid, model.VocabularySize,
                totalExamples, features);
            double difference = validScore - invalidScore;
            double probability;
            if (difference >= 0)
            {
                probability = 1.0
                    / (1.0 + Math.Exp(-Math.Min(difference, 700.0)));
            }
            else
            {
                double exponent =
                    Math.Exp(Math.Max(difference, -700.0));
                probability = exponent / (1.0 + exponent);
            }

            return new NlpNamePrediction
            {
                IsValid = probability >= 0.5,
                Probability = (float)probability
            };
        }

        private double CalculateScore(
            ClassStatistics statistics,
            int vocabularySize,
            int totalExamples,
            IEnumerable<string> features)
        {
            double score = Math.Log(
                (statistics.ExampleCount + 1.0)
                / (totalExamples + 2.0));
            double denominator = statistics.FeatureCount
                + vocabularySize;

            foreach (string feature in features)
            {
                int count = statistics.GetCount(feature);
                score += Math.Log((count + 1.0) / denominator);
            }

            return score;
        }

        private static IEnumerable<string> ExtractFeatures(
            string label, string shapeType)
        {
            string normalizedLabel = Regex.Replace(
                    label ?? string.Empty, @"\s+", " ")
                .Trim()
                .ToLowerInvariant();
            string normalizedType =
                (shapeType ?? string.Empty).Trim().ToLowerInvariant();

            yield return "type:" + normalizedType;
            yield return "length:"
                + Math.Min(12, normalizedLabel.Length / 5);

            MatchCollection matches = Regex.Matches(
                normalizedLabel, @"[\p{L}\p{Nd}]+");
            yield return "words:" + Math.Min(10, matches.Count);
            if (matches.Count == 0)
                yield return "word:<empty>";

            for (int index = 0; index < matches.Count; index++)
            {
                string word = matches[index].Value;
                yield return "word:" + word;
                if (index == 0)
                    yield return "first:" + word;
                if (index == matches.Count - 1)
                    yield return "last:" + word;
            }

            string padded = "^" + normalizedLabel + "$";
            for (int index = 0; index <= padded.Length - 3; index++)
                yield return "char3:" + padded.Substring(index, 3);
        }

        private static IList<NlpTrainingExample> ReadTrainingExamples()
        {
            Stream stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream(TrainingResourceName);
            if (stream == null)
            {
                throw new InvalidOperationException(
                    "The embedded NLP training data is unavailable.");
            }

            List<NlpTrainingExample> examples =
                new List<NlpTrainingExample>();
            using (stream)
            using (StreamReader reader = new StreamReader(stream))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] columns = line.Split('\t');
                    if (columns.Length != 3
                        || !bool.TryParse(columns[2], out bool isValid))
                    {
                        continue;
                    }

                    examples.Add(new NlpTrainingExample
                    {
                        Name = columns[0],
                        ShapeType = columns[1],
                        IsValidName = isValid
                    });
                }
            }

            if (examples.Count == 0)
            {
                throw new InvalidOperationException(
                    "The NLP training data contains no valid rows.");
            }

            return examples;
        }

        private sealed class ClassStatistics
        {
            private readonly Dictionary<string, int> featureCounts =
                new Dictionary<string, int>(StringComparer.Ordinal);

            public int ExampleCount { get; set; }
            public int FeatureCount { get; private set; }

            public void Add(string feature)
            {
                featureCounts.TryGetValue(feature, out int count);
                featureCounts[feature] = count + 1;
                FeatureCount++;
            }

            public int GetCount(string feature)
            {
                return featureCounts.TryGetValue(feature, out int count)
                    ? count
                    : 0;
            }
        }

        private sealed class BinaryClassModel
        {
            private readonly HashSet<string> vocabulary =
                new HashSet<string>(StringComparer.Ordinal);

            public ClassStatistics Valid { get; } =
                new ClassStatistics();
            public ClassStatistics Invalid { get; } =
                new ClassStatistics();
            public int VocabularySize =>
                Math.Max(1, vocabulary.Count);

            public void Add(
                bool isValid, IEnumerable<string> features)
            {
                ClassStatistics statistics =
                    isValid ? Valid : Invalid;
                statistics.ExampleCount++;
                foreach (string feature in features)
                {
                    statistics.Add(feature);
                    vocabulary.Add(feature);
                }
            }

            public void Validate()
            {
                if (Valid.ExampleCount == 0
                    || Invalid.ExampleCount == 0)
                {
                    throw new InvalidOperationException(
                        "Every shape type in the NLP training data must "
                        + "contain valid and invalid examples.");
                }
            }
        }
    }
}

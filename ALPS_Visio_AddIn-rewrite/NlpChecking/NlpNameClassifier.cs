using Microsoft.ML;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    /// <summary>
    /// Trains the PASS label classifier from the dataset bundled with the add-in.
    /// </summary>
    internal sealed class NlpNameClassifier
    {
        private const string TrainingResourceName =
            "ALPS_Visio_AddIn_rewrite.NlpChecking.training.tsv";

        private readonly object synchronization = new object();
        private readonly MLContext context = new MLContext(seed: 0);
        private PredictionEngine<NlpTrainingExample, NlpNamePrediction>
            predictionEngine;
        private int trainingExampleCount;

        public int EnsureTrained()
        {
            lock (synchronization)
            {
                if (predictionEngine == null)
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
                if (predictionEngine == null)
                    TrainCore();

                return predictionEngine.Predict(new NlpTrainingExample
                {
                    Name = label ?? string.Empty,
                    ShapeType = shapeType ?? string.Empty
                });
            }
        }

        private void TrainCore()
        {
            IList<NlpTrainingExample> examples = ReadTrainingExamples();
            IDataView trainingData = context.Data.LoadFromEnumerable(examples);

            IEstimator<ITransformer> pipeline = context.Transforms.Text
                .FeaturizeText("NameFeaturized",
                    nameof(NlpTrainingExample.Name))
                .Append(context.Transforms.Text.FeaturizeText(
                    "ShapeTypeFeaturized",
                    nameof(NlpTrainingExample.ShapeType)))
                .Append(context.Transforms.Concatenate(
                    "Features", "NameFeaturized", "ShapeTypeFeaturized"))
                .Append(context.BinaryClassification.Trainers
                    .SdcaLogisticRegression(
                        labelColumnName:
                            nameof(NlpTrainingExample.IsValidName),
                        featureColumnName: "Features"));

            ITransformer model = pipeline.Fit(trainingData);
            predictionEngine = context.Model.CreatePredictionEngine<
                NlpTrainingExample, NlpNamePrediction>(model);
            trainingExampleCount = examples.Count;
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
    }
}

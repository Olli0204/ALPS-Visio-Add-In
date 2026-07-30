# Third-Party Notices

## NLP PASS Checking

The NLP naming-check behavior and training data were adapted from:

- Project: `MatthesElstermann/NLPPASSCheckingBackup`
- Source: <https://github.com/MatthesElstermann/NLPPASSCheckingBackup>
- Revision reviewed: `f556b60e945d0b7894059100a83ee07c777b6c3b`
- Integrated data:
  `ALPS_Visio_AddIn-rewrite/NlpChecking/Resources/training.tsv`

No `LICENSE`, `COPYING`, or equivalent grant was present in the reviewed
revision. The repository owner's permission and the applicable terms must be
confirmed before redistributing the adapted implementation or training data.
This notice records provenance and is not itself a license grant.

The source project used Microsoft.ML. The integrated implementation replaces
that runtime dependency with a deterministic in-process text classifier so
retraining remains compatible with Visio's .NET Framework VSTO host.

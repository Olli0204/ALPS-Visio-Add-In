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

## ALPS Verification

The SID verification rules and two-model workflow were adapted from:

- Project: `andikra/ALPS-Verification-Thesis`
- Source: <https://github.com/andikra/ALPS-Verification-Thesis>
- Revision reviewed: `122e82c48fbb30bb995d258f184b29f593ebfef1`
- Copyright: Copyright (c) 2022 I2PM
- License: MIT
- Integrated model fixtures: `docs/verification-thesis/*.owl`

The console application itself was not copied. Its supported checks were
reworked as an offline, in-process verifier for the .NET Framework 4.8 VSTO
host, with structured results and explicit reporting of the prototype's
unimplemented scope.

MIT License

Copyright (c) 2022 I2PM

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## PASS to BPMN Converter

The BPMN model, conversion, layout, and serialization implementation was
adapted from:

- Project: `pass-bpmn-converter/pass-bpmn-converter`
- Source: <https://github.com/pass-bpmn-converter/pass-bpmn-converter>
- Revision reviewed: `27392ef8995b19c44cac3ba747a5ffd6abe04ca8`
- License: GNU General Public License v3.0

The source is included under
`ALPS_Visio_AddIn-rewrite/BpmnConversion/Upstream/`. Its complete license text
is retained in that directory. Compatibility-only changes adapt modern C#
constructs and collection APIs to the add-in's .NET Framework 4.8 host.
Redistribution of binaries containing this integrated source must comply with
the GNU General Public License v3.0.

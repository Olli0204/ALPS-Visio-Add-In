See the current [installation guide](docs/latex/main.pdf).

See [TODO.md](ALPS_Visio_AddIn-rewrite/TODO.md) for more information about the current state of the code. (German)

The **Verify ALPS Models** Ribbon command integrates the supported SID checks
from [andikra/ALPS-Verification-Thesis](https://github.com/andikra/ALPS-Verification-Thesis).
It compares an abstract specification OWL/RDF file with an implementation and
shows structured errors, warnings, and the remaining prototype scope.
The original thesis test models are documented in
[`docs/verification-thesis`](docs/verification-thesis/README.md).

The **Convert PASS to BPMN** Ribbon command integrates
[pass-bpmn-converter](https://github.com/pass-bpmn-converter/pass-bpmn-converter).
It converts a selected PASS/ALPS OWL or RDF model into a BPMN 2.0 file,
including generated diagram coordinates.

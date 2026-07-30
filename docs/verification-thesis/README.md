# ALPS Verification Thesis fixtures

These OWL files are exact copies of the model fixtures from:

- Repository: `andikra/ALPS-Verification-Thesis`
- Revision: `122e82c48fbb30bb995d258f184b29f593ebfef1`
- Original path: `alps .net api/ALPS Verification/Ont`
- License: MIT, Copyright (c) 2022 I2PM

The clearest specification/implementation pairs are:

| Specification | Implementation |
| --- | --- |
| `AbstractModel.owl` | `ImplementingModel.owl` |
| `Customer_is_king_Abstract.owl` | `Customer_is_king.owl` |
| `Customer_is_king_Abstract2.owl` | `Customer_is_king2.owl` |

`Impl.owl`, `Precendence.owl`, `Test.owl`, `Testmodel.owl`, and `VerTest.owl`
are retained with their original names as exploratory fixtures. In particular,
the spelling of `Precendence.owl` is preserved so provenance and upstream
comparisons remain unambiguous.

The four `Customer_is_king*.owl` files also retain the upstream default
namespace containing the literal segment `SID 18`. Strict XML namespace
validators report this unescaped space as a namespace warning. The files were
not silently normalized because this directory records the original fixtures;
this also makes parser-compatibility problems reproducible.

The upstream copies of `standard_PASS_ont_v_1.1.0.owl` and
`abstract-layered-pass-ont.owl` are deliberately not duplicated here. They are
ontology definitions rather than model fixtures, and equivalent ontologies
are already shipped under `ALPS_Visio_AddIn-rewrite/Resources`.

The two `[Test]_ALPS_Verification_*.owl` files one directory above are small,
add-in-owned regression fixtures with deterministic expected findings. They
are not part of the upstream repository.

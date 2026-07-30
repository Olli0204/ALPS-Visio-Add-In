# PASS to BPMN converter source

This directory contains the BPMN model, conversion, layout, and serialization
sources from `pass-bpmn-converter/pass-bpmn-converter`.

- Upstream revision: `27392ef8995b19c44cac3ba747a5ffd6abe04ca8`
- Source: <https://github.com/pass-bpmn-converter/pass-bpmn-converter>
- License: GNU General Public License v3.0; see `LICENSE` in this directory

Compatibility-only changes were made for the Visio add-in's .NET Framework 4.8
target: collection expressions and required members were replaced, and
`Queue.TryDequeue`/`Stack.TryPop` loops were rewritten for the framework API.

The integration also contains correctness fixes required by the add-in:

- model labels are kept as BPMN names instead of being emitted as invalid IDs;
- message flows and their diagram edges connect matching send/receive tasks;
- edge docking remains finite for horizontal and cyclic layouts;
- exclusive branches receive labeled, descriptive formal conditions;
- only supported timer, conditional, and error events become boundary events.

# First slide number: PowerPoint observation

On 2026-10-28, Microsoft PowerPoint for Mac 16.104 opened a two-slide copy produced by `attribute-v2` inserting `firstSlideNum=17` into `.tmp/native-field-powerpoint-01/edited.pptx`. Before opening, the generated file SHA-256 was `0819fd2c1e993831491c36194d38935915c17b7bf064b16db59d86428d2225a2`. The on-screen slide numbers and dynamic fields showed 17 and 18; neighboring text and date fields were visible.

In PowerPoint, the first slide was cut and pasted after the second, then the same disposable copy was saved, closed and reopened. The reordered first slide displayed `Page 1: 17 / Date: 28 October 2026`; the second displayed `EDIT Page 2: 18 / Date: 28 October 2026` before closing. Reopening showed the first slide and both thumbnail numbers 17 and 18. The saved `.tmp/first-slide-number-native-v2/input.pptx` has SHA-256 `363fb14042bdb22ca30daf00e0692d5a829b584d654db7a8a1d1c7d38d0505c5`; independent parser inspection reports two slides and an explicit starting number 17. The original generated file was overwritten by PowerPoint; its hash identifies it but its bytes were not retained separately. No screenshots were retained.

This single observation does not validate clearing, all values, additional applications, three reopen cycles, or unchanged field identities and all package parts after PowerPoint edits.

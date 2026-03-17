# VBA Regex Toolbox


| Action  | What it does  | Core members  | Minimal VBA code  |  |  |
| --- | --- | --- | --- | --- | --- |
| **Test (boolean)**  | Check if text matches a pattern (anywhere or exact via anchors) | `.Test(text)`  | `vba\nDim re As Object: Set re = Rx("^[A-Z]\\d{5}$")\nIf re.Test("A18370") Then Debug.Print "Match"\n`  |  |  |
| **Execute (iterate matches)**   | Get all matches and groups  | `.Execute(text)` → `Matches` → `Match.SubMatches`  | `vba\nDim re As Object, m As Object\nSet re = Rx("(\\d{4})-(\\d{2})-(\\d{2})", False, True)\nFor Each m In re.Execute("On 2025-10-31 and 2026-01-01")\n  Debug.Print m.Value, m.SubMatches(0), m.SubMatches(1), m.SubMatches(2)\nNext\n` |  |  |
| **Replace (transform text)**    | Substitute text based on a pattern                              | `.Replace(text, replacement)` (supports backrefs like `$1`) | `vba\nDim re As Object: Set re = Rx("(\\d{4})-(\\d{2})-(\\d{2})")\nDebug.Print re.Replace(\"2025-10-31\", \"$3/$2/$1\") '31/10/2025\n`  |  |  |
| **Capture groups**  | Extract specific subparts  | Parentheses `( ... )`, accessed via `Match.SubMatches(i)`   | `vba\nDim re As Object, m As Object\nSet re = Rx("([A-Z])(\\d{5})")\nSet m = re.Execute(\"A18370\")(0)\nDebug.Print m.SubMatches(0) 'A\nDebug.Print m.SubMatches(1) '18370\n`  |  |  |
| **Word boundaries & anchors**   | Constrain position (start/end/word)  | `^`, `$`, `\b`, `\B`  | `vba\nSet re = Rx(\"\\b[A-Z]\\d{5}\\b\")\nDebug.Print re.Test(\"Ref A18370 due\") 'True\n`  |  |  |
| **Lookarounds**  | Match with context without consuming it  | `(?=...)`, `(?!...)`, `(?<=...)`, `(?<!...)`  | `vba\n' Find ABC only when followed by -123\nSet re = Rx(\"ABC(?=-123)\")\nDebug.Print re.Test(\"ABC-123\") 'True\n`  |  |  |
| **Quantifiers**  | Control repetition  | `?`, `*`, `+`, `{m,n}`  | `vba\nSet re = Rx(\"[A-Z]{2}\\d{4,6}\")\nDebug.Print re.Test(\"AB183700\") 'True\n`  |  |  |
| **Character classes & escapes** | Sets and types  | `[A-Z] [0-9] \d \w \s`, negation `[^...]`  | `vba\nSet re = Rx(\"[A-Z]\\d{5}\")\n` |  |  |
| **Alternation & grouping**      | Choice between subpatterns  | \`A  | B`, grouping `( ... )\`  | \`\`\`vba\nSet re = Rx("(CAT | DOG)-\d+")\n\`\`\` |

---

[MOC](./tbx-00_MOC.md)

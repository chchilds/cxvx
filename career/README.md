# Christopher E. Childs — career documents (v15, November 2026)

Aligned resume, CV, executive bio, and commercial portfolio. Facts come from resume v14 ([PR #9](https://github.com/chchilds/cxvx/pull/9)), the public LinkedIn profile [linkedin.com/in/chrischildsraleigh](https://www.linkedin.com/in/chrischildsraleigh), and open commercial PRs in this repository. Customer TCV and pipeline figures from those PRs are not copied here.

## Documents

| Use | Markdown | Word | PDF |
| --- | --- | --- | --- |
| Resume (external) | [Resume - Chris Childs 2026 (v15).md](Resume%20-%20Chris%20Childs%202026%20(v15).md) | [docx](Resume%20-%20Chris%20Childs%202026%20(v15).docx) | [pdf](Resume%20-%20Chris%20Childs%202026%20(v15).pdf) |
| CV (longer) | [CV - Chris Childs 2026 (v15).md](CV%20-%20Chris%20Childs%202026%20(v15).md) | [docx](CV%20-%20Chris%20Childs%202026%20(v15).docx) | [pdf](CV%20-%20Chris%20Childs%202026%20(v15).pdf) |
| Executive bio | [Executive-Bio - Chris Childs 2026 (v15).md](Executive-Bio%20-%20Chris%20Childs%202026%20(v15).md) | [docx](Executive-Bio%20-%20Chris%20Childs%202026%20(v15).docx) | [pdf](Executive-Bio%20-%20Chris%20Childs%202026%20(v15).pdf) |
| Commercial portfolio | [Commercial-Portfolio - Chris Childs 2026 (v15).md](Commercial-Portfolio%20-%20Chris%20Childs%202026%20(v15).md) | [docx](Commercial-Portfolio%20-%20Chris%20Childs%202026%20(v15).docx) | [pdf](Commercial-Portfolio%20-%20Chris%20Childs%202026%20(v15).pdf) |

Source notes and open questions: [source-notes.md](source-notes.md).

## Rebuild

```bash
pip install -r career/requirements.txt
python3 career/build_documents.py
```

Edit `career/facts.py` and rebuild. Do not invent employment history.

# LaTeX CVs (Debasish Chakroborti)

Shared content files are the single source of truth. Website pages should match these facts.

## Build

```bash
# Academic (includes bibliography)
pdflatex cv-academic
# If `biber` fails on Apple Silicon lipo issues:
#   lipo "$(which biber)" -thin arm64 -output /tmp/biber-arm64 && /tmp/biber-arm64 cv-academic
biber cv-academic
pdflatex cv-academic
pdflatex cv-academic

# Industry
pdflatex cv-industry
pdflatex cv-industry
```

Copy PDFs to the site root for GitHub Pages:

```bash
cp cv-academic.pdf cv-industry.pdf ../
```

## Files

| File | Role |
|------|------|
| `cv-academic.tex` | Academic CV wrapper |
| `cv-industry.tex` | Industry CV wrapper |
| `objectives.tex` | Career objectives / research statement |
| `grants.tex` | Research grants and funding roles |
| `media.tex` | Media mentions and publicity |
| `talks.tex` | Talks, posters, and conference visits |
| `reviewing.tex` | Journal/conference reviewing and sub-reviewing |
| `collaborations.tex` | Research and industry collaborations |
| `mentoring.tex` | Student supervision and mentoring |
| `certifications.tex` | Professional certifications and PD |
| `employment.tex` | Shared employment |
| `education.tex` | Shared education |
| `skills.tex` | Shared skills |
| `awards.tex` | Shared awards / service |
| `teaching.tex` | Teaching (academic CV) |
| `projects.tex` | Selected projects (industry CV) |
| `publications.tex` | BibLaTeX publication list |
| `own-bib.bib` | Bibliography (synced from `research.md`) |
| `settings.sty` | Template style / biblatex setup |
| `photo.jpg` | Optional photo (enable `fullonly` in wrappers) |



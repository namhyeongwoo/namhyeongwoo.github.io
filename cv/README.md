# CV

`cv.json` is the single source of the CV. To update it:

```bash
cd cv
npm install            # first time only (installs docx)
node build.js          # writes ../files/CV_Hyeongwoo_Nam.docx
soffice --headless --convert-to pdf --outdir ../files ../files/CV_Hyeongwoo_Nam.docx
```

The website's CV button links to `files/CV_Hyeongwoo_Nam.pdf`. Keep `index.html` and
`cv.json` in sync when entries change.

# school-roster-pages

Static website for uploading an Excel file of students, grouping them by school, assigning students to day batches of 10, and exporting one PDF per school ordered by day.

The site includes two versions:

- English at `index.html`
- Armenian at `hy/index.html`

The PDF export is rendered from browser HTML instead of direct jsPDF table text, so Armenian, Latin, digits, and punctuation match the on-screen rendering for uploaded files too.

## Expected spreadsheet columns

- `studentid`
- `studentname`
- `school`

Optional:

- `day`

If `day` is missing, the site assigns students randomly within each school into `Day 1`, `Day 2`, and so on, using batches of 10 students per day.

## Local use

Open `index.html` in a browser and upload an `.xlsx` or `.xls` file.

The repository also includes `mock-students.xlsx` with:

- 100 total students
- 2 schools
- 50 students per school
- 5 days per school
- 10 students per day

## GitHub Pages deployment

This repo is configured to deploy to GitHub Pages using GitHub Actions on pushes to `main`.

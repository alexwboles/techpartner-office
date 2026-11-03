# Template gallery — competitor research (Nov 2026)

Compared before building, per the standing pre-build research rule. All
template _content_ in this repo is written originally for TechPartner Office;
nothing below copies a vendor's template text, layout, or artwork — only
category names and the shape of a library are compared.

## Microsoft 365 / Word (templates.office.com, File > New)

- Top-level browse: Business, Personal, Education, Events; search across
  thousands of templates. Very large library (hundreds per major category).
- Recurring categories: Resumes & Cover Letters, Reports, Invoices,
  Letters, Memos, Meeting Agendas & Minutes, Project Plans, Business Plans
  & Proposals, Calendars, Planners, Brochures, Flyers, Newsletters,
  Invitations, Business Cards, Labels.
- Pattern: one gallery across Word/Excel/PowerPoint, filtered by app;
  picking a template creates a new unsaved document from it.

## Google Docs / Sheets / Slides (Template gallery)

- Small curated gallery (~50 templates total), grouped Resumes, Letters,
  Personal, Work, Education.
- Docs: proposals, letters, newsletters, resumes. Sheets: budgets,
  calendars, attendance, invoices. Slides: photo albums, reports, pitches,
  proposals.
- Pattern: picking a template creates a new saved copy ("Copy of …") that
  opens immediately; the gallery is a start screen, not a store.

## WPS Office (Template Library)

- Very large library (marketed as 100,000+; a mix of local starters and
  online downloads, some behind Premium).
- Categories: Business Documents (resumes, business reports, project
  proposals, invoices, meeting minutes), Creative Presentations (corporate,
  academic, marketing pitches), Marketing Materials (brochures, flyers,
  posters), Personal Use (planners, invitations, greeting cards, journals).
- Pattern: templates are styled DOCX/PPTX/XLSX files with placeholders.

## LibreOffice (templates.libreoffice.org)

- Community library, organized by app (Writer / Calc / Impress) and by
  fine-grained categories: Accounting, Agenda, Budget, Business, Calendar,
  Card, Curriculum/Resume, Invoice, Letter, Memo, Newsletter, Presentation,
  Timesheet, To Do List, Education, and more.
- Pattern: download a template file, then File > New from it.

## Must-haves (appear in ALL four)

Resumes & CVs, cover letters, business letters, invoices, business plans
& proposals, meeting agendas & minutes, reports, budgets, schedules /
planners, presentations (pitch, sales, training, project status), flyers &
newsletters, forms & checklists, education (lesson plan, syllabus).

## What TechPartner Office ships (and how it differs)

- 52 original templates across 13 categories (see
  `apps/shell/src/shared/template-defs.ts`), covering every must-have above.
- Templates are _definitions_, not binary files: Docs templates are HTML
  fed through the app's own AI-document pipeline; Sheets templates are
  row/formula definitions built into a real .xlsx; Slides templates are
  slide definitions built into a real .pptx through pptx-engine. Every
  template opens as a normal editable document in the right app — no dead
  previews, no downloads, no account, nothing behind a paywall.
- Creation follows the app's own generated-document precedent (the same
  shape as AI create_document): the file lands in the selected folder, or
  the default save folder, under the template's title (unique-suffixed on
  clash) and opens as a tab. Docs templates additionally go through the
  pending-folder move, so a folder-view creation ends up in that folder.
- The gallery keeps the plain file picker reachable ("Open local file"),
  so the old From-template picker behavior is not lost.

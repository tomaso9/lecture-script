---
name: lecture-script
description: Turn a recorded class (transcript + slide deck) into a tight, engaging, studio-ready video lecture script as a color-coded Word document, following the CODL writing and evergreen guides. Use when an instructor wants to re-record a live class for online delivery.
---

# Lecture script

Inputs: a class transcript (usually auto-captioned and messy) and the slide
deck. Output: a Word script the instructor reads in a studio, split into
self-contained segments, with every sentence color-coded by provenance.

The helper for all file work is `scripts/lecture_tools.py` (needs Python 3
with `python-docx` and `python-pptx`). Never hand-format the Word file.

## Rules that apply throughout

**Two authorities.** `references/Video Lecture Script - writing guide v1.docx`
and `references/Video Lecture Script - evergreen v1.docx`, read with
`lecture_tools.py docx` at the start of every run. If a newer version of
either guide is attached in the conversation, the attached one wins for that
run; tell the user the copy in `references/` is out of date.

**Length.** Total spoken time for a module is 20 minutes by default, in
segments of 5 to 15 minutes each. State the default before proposing an
outline and let the user change it. The script is a condensed version of the
class, not the class reflowed: the slides set the spine, and the transcript
supplies the words.

**Provenance, sentence by sentence.**
- Black: the instructor's own words from the transcript. Repairing a
  caption, dropping filler, or completing a broken sentence keeps it black.
- Dark red: anything newly written, including recomposed wording, slide text
  spoken aloud (the instructor wrote it but never said it), and speaker notes.
- When in doubt, red. Never silently improve the instructor's words.

**Sound like the instructor, not like an AI.**
- Select the instructor's sentences before writing new ones. New writing is
  for hooks, transitions, and gaps the slides show the transcript missed. A
  segment over about a third red is a signal to go back to the transcript.
- The transcript is the only source of the instructor's voice. Before
  drafting, write a short voice sketch from it: how they open a point, how
  they ask questions, recurring phrases, favorite examples, typical sentence
  length, what they do when a student pushes back. Every red line is written
  against that sketch and must sit next to black lines without a visible
  seam.
- Drop verbal tics ("right?", "okay so") but keep one or two of the
  instructor's classroom check-ins per module, in their own words and where
  they actually said them, such as "Does that make sense?" They make the
  video feel like a class rather than a broadcast.
- If `references/style-sample.docx` exists, read it as a model of what a
  finished script looks like: structure, hook, cue style, pacing, how much
  gets rewritten. It is not a voice model. Do not borrow its phrasing.
- Banned tells in red lines: "let's dive in", "delve", "it's important to
  note", "think of it as", "in this segment we will", "today we're going
  to", tidy three-item lists for their own sake, a summary at the end, a
  string of rhetorical questions, "not only X but Y", "a testament to",
  "navigate", "landscape", "crucial", "robust". Rewrite before the user sees
  them.
- Optional voice module: if `references/voice-oliver.md` exists and the user
  did not say "plain voice", apply it to red lines only, within its limits.

## Workflow

1. **Read everything.** Dump both guides and the style sample with
   `lecture_tools.py docx`. Dump the deck with `lecture_tools.py pptx`
   (slide text, notes, diagram text, hidden slides marked). Read the
   transcript. Check they are the same session: slide titles should appear in
   the talk. If they do not, say so and stop. If a slide is an image the dump
   cannot describe, ask the user what it shows rather than guessing.

2. **Propose an outline and wait.** Before writing any script text, show:
   - the themes the deck carries, which ones get priority and which get cut,
     with a one-line reason each;
   - content that will not be recorded: activities, logistics, due dates,
     reading assignments, student discussion, anything tied to a week or
     module number;
   - the voice sketch drawn from the transcript, in five to eight lines;
   - the segments: title, central question, slides covered, target minutes,
     and the transcript passages that anchor each (a few words each);
   - the total against the length default.
   Wait for approval. Do not draft until the user says go.

3. **Draft as a plain-text script** in the format below, one sentence per
   line, saved next to the inputs as `<lecture>-script.txt` (the builder
   removes it at the end). Write against the
   writing guide's checklist: a hook that puts the student in a situation,
   one central question per segment stated early, concrete anchors, varied
   sentence length with at least one very short sentence, direct address, a
   spoken description of anything important shown on screen, an ending that
   leaves a question or tension rather than a summary. Segments are
   self-contained: no "last week", "next video", or module numbers, in the
   narration or on any slide cued.

4. **Check and fix.** Read the draft against both guides once more. Apply
   the evergreen checklist as a separate pass (relative dates, titles that
   change, "current", course logistics, back-references). Fix what you find.
   Report only the misses that need the instructor's judgement, in a few
   lines.

5. **Build and report.** Run
   `python scripts/lecture_tools.py build <lecture>-script.txt` and relay its
   per-segment word and minute table verbatim, including any flags. Fix
   flagged lengths and rebuild. When the numbers are right, run it once more
   with `--rm` so only the `.docx` is left next to the inputs, and tell the
   user where it is. They edit the Word file directly from there.

## Script text format

```
Title: What is policy design?
# Segment 1: Why does a policy look the way it looks?
[Open on camera, medium shot. No slide for the first lines.]
Read the Paris Climate Agreement closely and three things jump out.
+ Every one of those features was put there on purpose.
## One: administrators are designers
[Slide 6: three-part graphic. The next three lines describe it aloud.]
// a comment for yourself; ignored by the builder
```

Plain line: instructor's words (black). `+` prefix: new writing (dark red).
Whole line in `[ ]`: production cue (gray italic, not read aloud). `#`
starts a segment, `##` a sub-heading inside one. Cues name the slide number
and say what is on it. The builder adds the color key and per-segment counts.

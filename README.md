# lecture-script

A Claude Code skill that turns a recorded live class into a studio-ready
video lecture script. Give it the class transcript and the slide deck; it
proposes an outline for your approval, drafts a script in your own words
where possible, checks it against the Syracuse CODL writing and evergreen
guides, and builds a color-coded Word document:

- **Black** is what you said in class, lightly repaired.
- **Dark red** is new writing you did not say.
- **Gray italic** is a production cue, not read aloud.

The whole module targets 20 minutes by default, in segments of 5 to 15
minutes. You can change both at the start of a run.

## Install

Clone into your Claude Code skills folder and install two Python packages:

```
git clone <this repo> ~/.claude/skills/lecture-script
pip install python-docx python-pptx
```

On Windows the skills folder is `C:\Users\<you>\.claude\skills\`.

## Use

In Claude Code, from the folder holding your transcript and deck:

```
/lecture-script  transcript.txt  "My deck.pptx"
```

Claude reads both, shows an outline, waits for your approval, then drafts,
checks, and builds `<lecture>-script.txt` and `<lecture>-script.docx` next
to your inputs. Edit the text file and rebuild any time:

```
python ~/.claude/skills/lecture-script/scripts/lecture_tools.py build my-lecture-script.txt
```

## Voice

The skill writes new lines to match the instructor's voice, inferred from the
transcript and from `references/style-sample.docx` if present. Replace that
file with a script of your own that you like.

`references/voice-oliver.md` adds a light "Last Week Tonight" register to new
lines: deadpan about absurd facts, escalation, a provocative close, one or two
beats per segment at most. It is optional. Delete the file to turn it off
permanently, or say "plain voice" when invoking the skill to skip it once.

## What is in the box

| Path | What |
|---|---|
| `SKILL.md` | The workflow and rules Claude follows |
| `references/Video Lecture Script - writing guide v1.docx` | CODL writing guide, read at the start of every run |
| `references/Video Lecture Script - evergreen v1.docx` | CODL evergreen checklist, same |
| `references/style-sample.docx` | An approved script used as a register model (optional) |
| `references/voice-oliver.md` | Optional voice module (delete to disable) |
| `scripts/lecture_tools.py` | Dumps pptx and docx to text; builds the Word script |

The two CODL guides are authored by the Syracuse University Center for Online
and Digital Learning. They are included here for use by the skill; do not
redistribute them outside that purpose.

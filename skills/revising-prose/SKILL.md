---
name: revising-prose
description: 'Revise a document for redundancy, conciseness, clarity, and a human voice using Richard Lanham''s "Revising Prose" (the Paramedic Method and the Lard Factor). Use when: asked to edit, tighten, trim, shorten, cut word count, find repeated or redundant sections, proofread for wordiness, simplify, de-jargon, improve readability or flow, or make AI/LLM-generated text sound human-written. Fixes bureaucratic "Official Style" prose, passive voice, nominalizations, prepositional-phrase strings, "is" verbs, noun stacks, slow windups, redundant phrases, points repeated across sections, empty hedges, double negatives, and LLM tells (delve, robust, pinned, "lives in", "holds no", "with no", em-dash overuse, "not just X, but Y", trailing ", ensuring..." clauses). Works on Markdown, plain text, emails, self-evaluations, design docs, READMEs, and reports. Edits in place and reports the Lard Factor.'
argument-hint: 'File to revise; optionally the audience, a target word count, or "review only"'
---

# Revising Prose

Cut the lard, restore the action, keep every idea. This skill applies Richard A. Lanham's *Revising Prose*. The book diagnoses the **Official Style**: prose built on nouns instead of verbs, strings of prepositional phrases, "is" verbs, the passive voice, and slow windups. Its cure is the **Paramedic Method**, and it measures the result with the **Lard Factor**. The skill also removes the habits that make text sound machine-written, so documents drafted or edited by an LLM read as if a person wrote them.

## Ground Rules

1. **Cut words, never meaning.** Keep every fact, number, name, date, commitment, caveat, and technical term. A high Lard Factor that comes from deleting content is a failure.
2. **Default mode: edit in place, then report.** If the user says "review only", "suggest", or "don't edit", leave the file alone and report the proposed rewrites.
3. **Leave non-prose alone:** code blocks, inline code, commands, URLs, file paths, identifiers, direct quotations, citations, legal or contractual language, and required template text such as form prompts and mandated headings.
4. **Keep the structure:** Markdown syntax, heading levels, list and table layout, front matter, links, and anchors that other documents may reference.
5. **Keep the author's voice:** person (I/we/you), tone, and register. The result should sound like the author on a good day, not like you.
6. **Clarity beats brevity.** If a longer phrasing is clearer, use it. Unpacking a noun stack often *adds* words, and that's fine.

## Key Concepts

| Concept | Meaning |
|---------|---------|
| **Official Style** | Noun-heavy, passive, "is"-based prose with strings of prepositional phrases. It sounds impressive and says little. |
| **Lard** | Words that take up space without carrying meaning. |
| **Lard Factor** | `LF = (original words - revised words) / original words x 100%`. Lanham calls an LF of one-third to one-half normal when you first revise Official Style prose. |
| **"Who's kicking who?"** | Each sentence should show an actor doing an action, and a simple active verb should carry that action. |
| **Action** | The real verb that often hides inside a noun ("make a *decision*") or an adjective ("is *dependent* on"). |
| **LLM tells** | Words, metaphors, sentence patterns, and punctuation habits that make text sound machine-written. One is harmless; clusters give the text away. |

## Tool: `prose_audit.py`

[scripts/prose_audit.py](./scripts/prose_audit.py) needs Python 3.8+ and only the standard library. It counts words the same way every time, which makes the LF accurate. It skips code blocks, inline code, URLs, HTML, and front matter. **Never count words by hand.**

```
python <this-skill-dir>/scripts/prose_audit.py scan <file> [--top N]
python <this-skill-dir>/scripts/prose_audit.py compare <original> <revised>
python <this-skill-dir>/scripts/prose_audit.py dupes <file> [--top N] [--scope auto|section|block] [--ignore-section REGEX] [--level N]
```

- `scan` prints summary metrics and the highest-scoring sentences in Paramedic markup: `[prepositions]`, `{"to be" forms}`, `<nominalizations>`. It also lists prepositional strings, lard phrases with suggested replacements, double-negative candidates, There-is/It-is openers, long sentences, monotone paragraphs, LLM tells, paragraphs with 2+ em dashes, and arrows or emoji.
- `compare` prints the Lard Factor and a before/after metrics table in Markdown, ready to paste into the report. The table includes "Sentences repeated elsewhere" and "Words in duplicate sentences", so it shows whether the cuts removed repeats.
- `dupes` finds repeats across the whole document, which `scan` can't see because it reads one sentence at a time. It lists groups of sentences that make the same point, passages whose wording recurs in other sections (naming those sections), and a per-section echo table. It compares across sections when the document has three or more headings and across paragraphs otherwise; `--scope` overrides that. Words inside inline code count, because identifiers carry the topic in technical writing. Use `--ignore-section` (a regex, repeatable) to skip sections that repeat the body on purpose, such as a changelog.
- Flags are **candidates, not verdicts**. Judge each one in context; see [Guardrails](#guardrails-when-not-to-cut).

## Procedure

### 1. Read and frame

- Read the whole document before changing anything.
- Identify the audience, purpose, genre, and constraints: word limit, template, style guide, required sections.
- Name the reader's main question and the writer's main point. Every later cut serves these two.
- Infer what you can. Ask the user only when a wrong guess would cause real harm, such as being unsure whether a passage is mandated text.

### 2. Take a baseline

- Copy the original to the system temp directory, e.g. `%TEMP%\revising-prose\<name>.orig.md` on Windows or `$TMPDIR/revising-prose/` elsewhere. Never put this copy in the workspace. It serves as the LF baseline and as the user's undo.
- Run `scan` on the original. Note the summary metrics and the worst sentences.
- For a document longer than a few pages, run `dupes` too and keep its output for step 3.

### 3. Document pass: cut large redundancies first

Make the big cuts first so you don't polish sentences you will later delete.

- Find points made more than once (intro, body, conclusion, summary bullets). Keep the strongest version and merge any unique details into it. In a long document, start from `dupes`: each group names sentences that make the same point, and each echoed passage names where else its wording appears.
- For each group, decide which section owns the point. Keep the best wording there, and cut the other copies or reduce them to a pointer ("see section 6.2"). The "up to N words" figure assumes you keep only the longest copy.
- Apply the clear-cut cuts, such as a sentence that repeats another nearly word for word. Offer the structural ones as options in the report, with the words each would save: a summary table that restates the sections it summarizes, or a threat-model row that restates a mitigation. Removing those changes how the document reads.
- Leave deliberate repeats alone: summaries and abstracts, changelogs, test plans that restate requirements, coverage tables, glossary entries, and prerequisites repeated in how-to steps. Ask the author when unsure, or re-run `dupes` with `--ignore-section` to hide them.
- Cut sentences that only restate the heading ("This section describes the build process.").
- Cut signposts and recaps that repeat content ("As mentioned above", "In summary"). Keep a summary only when a long document needs one for skimmers.
- Merge duplicate or overlapping bullets.
- Put the main point first in each section and paragraph.

### 4. Sentence pass: the Paramedic Method

Work through every sentence. Start with the worst ones from `scan`, but don't stop there. The full guide with worked examples is in [references/paramedic-method.md](./references/paramedic-method.md).

1. **Circle the prepositions.** Three or more in a row ("the review *of* changes *to* the plan *for* the team") signal a missing verb.
2. **Circle the "is" forms** (is, are, was, were, be, been, being). They often stand in for a buried action or mark a passive.
3. **Find the action. Ask "who's kicking who?"** Identify the actor and what the actor does.
4. **Put the action in a simple active verb,** with the actor as the subject: "make a decision" becomes "decide", and "is dependent on" becomes "depends on".
5. **Start fast.** Cut slow windups such as "It is important to note that", "There are several factors that", and "The purpose of this section is to".
6. **Read the passage aloud with emphasis and feeling.** Wherever you stumble, run out of breath, or can't find the stress, rewrite.
7. **Mark the rhythmic units** of each long sentence with `/`. Each unit should be a natural phrase, and the sentence should build toward its key word.
8. **Mark the sentence lengths** across a paragraph with `/`. Vary them.

### 5. Lard pass

Sweep up the remaining lard with [references/lard-catalog.md](./references/lard-catalog.md):

- compound prepositions and verb phrases → single words ("in order to" → "to", "prior to" → "before")
- redundant pairs and modifiers ("each and every", "past history", "completely eliminate")
- redundant categories ("in a timely manner" → "promptly", "large in size" → "large")
- empty intensifiers and reflexive hedges ("very", "really", "basically", "I think")
- **double and stacked negatives → positive statements** ("not uncommon" → "common"; "cannot start without a key" → "needs a key to start")
- noun stacks → unpacked phrases ("firmware update delivery mechanism" → "how firmware updates are delivered")
- jargon and buzzwords → plain words ("leverage" → "use", "utilize" → "use")

### 6. Paragraph pass: rhythm and engagement

- **Vary sentence length.** Break up runs of similar lengths. A short sentence after a long one adds punch.
- **End sentences on the important word.** Readers stress the end of a sentence, so put the new or key information there.
- **Use concrete subjects:** people and things doing things. Replace abstractions with the examples and numbers the document already supports. **Never invent facts.**
- **Keep lists and series parallel** ("design, build, and test", not "design, building, and the testing phase").
- **Keep transitions that show logic** (because, but, so, instead). Cut ones that only fill space.
- **Split sentences over about 30 words** unless their structure earns the length. Combine choppy runs of very short sentences.
- **Keep it human.** Use "you" and "we" where the genre allows. Prefer the familiar word to the impressive one.

### 7. Human-voice pass: remove LLM tells

Apply this pass to every document, whoever drafted it. Work from the "LLM tells" section of `scan` and from [references/llm-tells.md](./references/llm-tells.md). Fix paragraphs with several tells first.

- **Things that come alive:** Code, keys, and data don't live, sit, own, want, see, or travel. "Afterwards lives only in protected memory" becomes "afterwards exists only in protected memory".
- **Vivid stand-ins for precise terms:** Use the field's own term. In code signing, a key is trusted, authorized, enrolled, or provisioned, not "pinned". Data is sealed to a PCR, and a measurement is extended into one.
- **LLM-favored words:** delve, robust, seamless, crucial, pivotal, comprehensive, intricate, landscape, testament, underscore, foster.
- **Sentence patterns:** "not just X, it's Y"; trailing ", ensuring..." clauses; "The result?" setups; reflexive groups of three; vague "This ensures..."; signpost openers and summary closers.
- **Punctuation and formatting:** about one em dash per paragraph at most; no arrows or emoji in prose; bold only for defined terms and warnings; prose instead of bullets when the ideas connect.
- **Content:** generic claims where the author would give specifics, uniform confidence, and actions without actors. **Never invent the specifics**; flag them for the author.
- **Over-compression:** conditions hidden in noun phrases ("only in a boot in which..."), modifiers on the wrong noun ("a deployed, secure boot"), and events as actors ("A boot ... updates nothing"). Unpack them even when that adds words.
- **Negated nouns:** "holds no secrets" becomes "doesn't hold any secrets", "with no user interaction" becomes "without user interaction", and "shall contain no" becomes "shall not contain". Negate the verb, not the noun.

Don't swap one LLM word for another ("delve" → "explore", "robust" → "resilient"). Don't over-correct either: an occasional em dash or list of three is normal.

### 8. Verify

- **Meaning:** Compare the original and the revision section by section. Every claim, number, name, date, caveat, and commitment must survive, and nothing new may appear. If you're unsure whether an edit changed the meaning, list it for the author instead of guessing.
- **Small words that drift:** Even published model revisions slip here, so check each one:
  - quantifiers (all / most / some)
  - time qualifiers ("in most years" is not "in recent years")
  - modals (must / should / may)
  - attributions ("It is widely known that X" claims that others agree; plain "X" makes the author the source)
- **Measure:** Run `compare <temp-copy> <file>`.
  - Low LF on Official Style text (under about 10%) means you should do another pass.
  - High LF (over about 50%) means you should check the meaning again before reporting.
- **Repeats:** After a redundancy pass, "Sentences repeated elsewhere" and "Words in duplicate sentences" should fall. Re-run `dupes` to confirm that the groups you meant to cut are gone.
- **Re-scan:** Run `scan` on the revision. Fix every remaining "LLM tells" hit, or keep it on purpose and list it under "Please review" with the reason (for example, "no later than" is a spec limit).
- **Format:** Make sure the Markdown still renders: headings, lists, tables, links, and code fences are intact. Check for editor errors or lint warnings if available.

### 9. Report

Use this template:

```markdown
## Revision summary: <file>

**Lard Factor: 34.5%** (1,240 → 812 words)

<metrics table from `compare`>

### What changed
- 3–6 bullets naming the patterns fixed, with counts where useful
  (e.g., "Turned 14 'the X of' nominalizations into verbs"; "Replaced 6 repeated passages with pointers,
  cutting repeated sentences from 154 to 142").

### Before → after
2–4 of the most instructive sentence rewrites, each with its own LF.

### Please review
- Edits where the meaning may have shifted, and open questions for the author, with links to lines.
- Structural repeats offered as options and not applied, each with the words it would save.

Original saved at: <temp path>
```

## Guardrails: When Not to Cut

- **Passive voice is correct** when the actor is unknown, unimportant, or obvious, or when the object is the topic ("The fix was backported to the 2024 release").
- **Keep "is" for identity and definition** ("Latency is the time between request and response").
- **Keep established technical nouns** such as "authentication", "memory allocation", and "integration test". They are terms, not lard.
- **Keep hedges that report real uncertainty** ("likely", "in our tests", "as of March"). Cut only the reflexive ones.
- **Repeat defined terms.** Don't swap in synonyms to avoid repetition. In technical writing, consistency beats variety.
- **Keep repeats the genre needs.** Executive summaries, changelogs, per-section prerequisites in a how-to, and test plans that restate requirements are meant to repeat. `dupes` flags them. Cut them only if the author agrees.
- **Keep terms of art that look like LLM tells:** certificate pinning, pinned memory, pinned dependency versions, GPIO pins, attack surface, trust anchor, clock and power gating, the PI spec's Hand-Off Blocks, and any jargon the document's audience uses every day.
- **Don't write telegraphese.** Articles, conjunctions, and logical transitions help the reader.
- **Use understatement ("not bad") only on purpose,** where the genre allows it. That's rare in technical and business writing.
- **Keep required wording** in forms such as self-evaluations, compliance documents, and templates. Revise the author's answers, not the prompts.
- **Don't chase the number.** LF measures lard removed, not quality.

## Specifications and Requirements

When the document uses normative keywords (shall, must, required), precision outranks brevity. `scan` reports the keyword count. Apply these rules on top of the procedure:

- **Keep normative keywords exactly as written.** Don't change shall, must, should, or may, or their negatives, into each other or into "will". They carry defined meanings (RFC 2119/8174 or the spec's own conventions).
- **Write conditions as conditions:** use if, only if, when, or unless, not a noun with an "in which" clause.
- **Fold a separate negative rule into the conditions.** "Release K only if A and B. If C happened, don't release K." means "Release K only if A, B, and C didn't happen." One requirement is shorter and can't be read as two independent rules.
- **Negate the verb, not the noun:** "shall not contain executable code", not "shall contain no executable code"; "only if the UEFI Shell wasn't loaded", not "only if no UEFI Shell was loaded". Use contractions where the author does, but never contract normative keywords (shall not, must not).
- **One condition per clause, each with its real subject.** "The platform is in Deployed Mode" and "Secure Boot is enabled" are separate conditions, not "a deployed, secure boot".
- **Prefer a list for three or more conditions:** "only if:" followed by the items, with "; and" (AND) or "; or" (OR) before the last one.
  - Use one sentence instead when a list doesn't fit. Examples: the requirement already sits in a list item and nesting would clutter it, it's in a table cell, or the document's format avoids lists.
  - In the sentence, put any condition introduced by a reporting verb ("BDS reports that...") last, so the verb can't seem to cover the conditions before it.
  - Put each section reference right after the condition it supports.
- **Name the responsible component** at least once per requirement.
- **Repeat the defined term** (db, PK, KEK) instead of "it", "the database", or "nothing".
- **Make scope explicit:** which boot, which phase, until when.
- **Keep one authoritative statement per requirement.** When two sections state the same requirement, keep the normative sentence in the section that owns it and point to it from the other. Each removed `shall` or `must` lowers the keyword count that `scan` reports, so mention the removal in the report. Don't pick between statements that disagree; ask the author.
- **Ask the author about ambiguities you uncover** instead of resolving them yourself.

See [references/llm-tells.md](./references/llm-tells.md) section 7 for over-compressed specification text.

## LF Expectations

| Starting text | Typical LF |
|---------------|------------|
| Unrevised bureaucratic / Official Style | 33–50% (Lanham's norm for a first revision) |
| Ordinary first draft | roughly 15–30% |
| Already edited prose | roughly 5–15% |
| Over-compressed or normative text (specs, requirements, LLM drafts) | can be negative; precision first |

If the LF is below 5%, say the document was already lean, and focus the report on clarity and rhythm instead. If it is negative because you unpacked compressed logic, say so; that's a good result.

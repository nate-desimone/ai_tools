# The Paramedic Method in Detail

Richard Lanham calls this method "paramedic" because it is emergency care for prose in the Official Style. It won't create a style for you, but it reliably revives sentences that drown in nouns, prepositions, and "is" verbs. The eight steps below follow the later editions of *Revising Prose*. The [Purdue OWL handout](https://owl.purdue.edu/owl/general_writing/academic_writing/paramedic_method.html), adapted by Richard Johnson-Sheehan, teaches a shorter seven-step version. It drops Lanham's sound-and-rhythm steps (6–8) and adds two explicit steps: "move the doer into the subject" and "eliminate any redundancies". This skill folds those two into steps 3–4 and the lard pass.

All examples here are original. Word counts and LFs come from `prose_audit.py`.

---

## Step 1: Circle the Prepositions

**Look for:** of, in, to, for, by, with, on, at, from, about, into, through, upon, prior to, with regard to, and similar words.

**Why:** In the Official Style, prepositional phrases carry the meaning that verbs should carry. A string of three or more ("the *review of* changes *to* the plan *for* the team") is the clearest warning sign. A useful rule of thumb: if more than about one word in eight is a preposition, the sentence probably hides a verb.

**`of` deserves special suspicion.** "The X of the Y" is usually a verb in disguise:

| Before | After |
|--------|-------|
| the implementation of the policy | implementing the policy / we implemented the policy |
| the failure of the test | the test failed |
| a reduction of costs | reduce costs / costs fell |

## Step 2: Circle the "Is" Forms

**Look for:** is, are, was, were, be, been, being, am, and contractions such as it's, there's, and they're.

**Why:** "Is" links things; it doesn't *do* anything. When "is" is a sentence's main verb, the real action usually hides in a noun or adjective nearby. "Is" + past participle marks a passive, which often hides the actor.

| Pattern | Example | Fix |
|---------|---------|-----|
| is + adjective + preposition | The schedule *is dependent on* the vendor. | The schedule *depends on* the vendor. |
| is + noun + of | This *is a reflection of* our priorities. | This *reflects* our priorities. |
| passive with known actor | The design *was reviewed by* the architects. | The architects *reviewed* the design. |
| There is / are ... that | *There are* three bugs *that* block the release. | Three bugs block the release. |
| It is ... that / to | *It is* necessary *to* rebuild the cache. | Rebuild the cache. / We must rebuild the cache. |

## Step 3: Find the Action. Ask "Who's Kicking Who?"

Ask two questions of every sentence:

1. **What happens?** That's the action. Look inside nouns ending in -tion, -ment, -ance, -ence, -ity, and -ness, and inside "is + adjective" phrases.
2. **Who does it?** That's the actor. It may be buried in a "by" phrase or an "of" phrase, or it may be missing entirely.

If the actor is missing, decide whether the reader needs to know it. In a status report, "the build was broken" hides responsibility. In a description of what a system does, "the image is signed" is often fine.

If the actor is missing but known, supply it. In reports and proposals, the hidden actor is often "we" or "I": "An investigation into the outage is underway" becomes "We are investigating the outage." Never invent an actor the document doesn't identify.

## Step 4: Put the Action in a Simple Active Verb

Make the actor the subject and the action the verb. Check both halves. "An analysis of the failure was performed" → "The failure was analyzed" fixes the verb but still hides the doer. "The team analyzed the failure" fixes both.

"Simple" means a single verb, not a compound phrase built around a noun:

| Compound (weak) | Simple (strong) |
|-----------------|-----------------|
| make a decision | decide |
| perform an analysis of | analyze |
| give consideration to | consider |
| provide assistance to | help |
| is in agreement with | agrees with |
| has an impact on | affects |
| conduct a review of | review |
| achieve a reduction in | reduce |
| serves to illustrate | illustrates |

## Step 5: Start Fast

Readers decide within a few words whether a sentence is worth their attention, so don't spend those words clearing your throat.

**Slow windups to cut:**

- "It is important to note that ..." / "It should be noted that ..."
- "The purpose of this document is to ..." / "In this section, we will discuss ..."
- "There are a number of reasons why ..."
- "In today's fast-paced environment, ..."
- "When it comes to X, ..." / "With regard to X, ..."
- "The fact of the matter is that ..."
- "As previously mentioned, ..."

Start with the subject and verb. If a sentence opens with a long introductory phrase, ask whether that phrase belongs at the end or nowhere at all.

## Step 6: Read Aloud with Emphasis and Feeling

You can't hear the text, so simulate a reader:

- **Breath test:** Can a reader say each sentence, or each unit between commas, in one breath? About 25–30 words without a pause is the limit.
- **Stress test:** Where does the voice naturally land? The stress should fall on the most important word, usually near the end. If the sentence ends on a filler phrase ("... in this area", "... going forward", "... as well"), move or cut it.
- **Stumble test:** Re-reading a phrase to parse it means it's broken. Typical culprits are noun stacks, nested clauses, and stacked negatives.
- **Sound test:** Repeated sounds ("the implementation of the configuration of the validation") and jingles ("ensure the assurance") distract the reader.

## Step 7: Mark the Rhythmic Units with "/"

For each long sentence, put a `/` between natural phrase units:

> The committee recommends / that the security team review firmware changes / before release.

Check that:

- each unit is a meaningful phrase, not a fragment like "of the".
- the units don't all have the same shape and length.
- the sentence builds toward its key unit instead of trailing off.
- each unit has a job. If you can't say what a unit adds, cut it.

## Step 8: Mark the Sentence Lengths with "/"

Write the word count of each sentence in a paragraph: `18 / 21 / 19 / 22`. Similar numbers mean monotone prose, and monotone prose puts readers to sleep. `scan` lists such paragraphs under "Rhythm".

Fix monotony by:

- adding a short sentence for emphasis after a long one.
- combining two related short sentences with a subordinate clause.
- splitting a sentence over 30 words at its natural joint.
- varying how sentences open: subject, introductory phrase, or dependent clause.

**Before** (7 / 8 / 9 / 9: monotone, list-like):

> The build system compiles each module separately. The linker combines the modules into one image. The signing tool adds a signature to the image. The release script uploads the image to the server.

**After** (7 / 21):

> The build system compiles each module separately. The linker then combines them into one image, which the signing tool signs and the release script uploads to the server.

---

## Worked Examples

### Example 1: The Full Method

**Before (36 words):**

> It is the recommendation of this committee that consideration be given to the implementation of a process for the review of firmware changes by members of the security team prior to the release of the product.

**Steps 1–2 (markup):**

> It {is} the \<recommendation> [of] this committee that \<consideration> {be} given [to] the \<implementation> [of] a process [for] the review [of] firmware changes [by] members [of] the security team prior [to] the release [of] the product.

That's 9 prepositions (one word in four), 2 "is" forms, an 8-preposition string, and 3 nominalizations.

**Step 3, who's kicking who?** The committee recommends. The security team reviews firmware changes before release.

**Step 4, simple verbs:** recommends, review.

**Step 5, start fast:** Replace "It is the recommendation of this committee that" with "The committee recommends".

**After (12 words, LF 67%):**

> The committee recommends that the security team review firmware changes before release.

### Example 2: There Is ... That

**Before (18 words):**

> There are a number of factors that are contributing to the delay in the delivery of the build.

**After (6 words, LF 67%):**

> Several factors are delaying the build.

Better still, if the document names the factors elsewhere, say what they are here.

### Example 3: Nominalization

**Before (8 words):** The team performed an analysis of the failure.

**After (5 words, LF 38%):** The team analyzed the failure.

### Example 4: Passive with a Known Actor

**Before (10 words):** The decision was made by management to delay the launch.

**After (6 words, LF 40%):** Management decided to delay the launch.

### Example 5: Double Negatives

**Before (9 words):** It is not uncommon for the test to fail.

**After (4 words, LF 56%):** The test often fails.

**Before (9 words):** The driver will not load without a valid signature.

**After (8 words, LF 11%):** The driver needs a valid signature to load.

The second fix saves only one word, but the reader no longer has to cancel out two negatives in their head. Clarity is the goal; word count is a symptom.

### Example 6: Noun Stack

**Before (11 words):** The platform firmware security vulnerability response process improvement plan is attached.

**After (15 words, LF −36%):** Attached is our plan to improve how we respond to security vulnerabilities in platform firmware.

The revision is *longer* and much clearer. A noun stack forces the reader to parse a chain of nouns backward. When a stack has four or more nouns, unpack it with a verb and a preposition or two, even if that costs words.

### Example 7: Slow Windup Plus Lard

**Before (31 words):**

> In today's rapidly evolving landscape, it is more important than ever before that each and every engineer on the team makes a concerted effort to write documentation in a timely manner.

**After (9 words, LF 71%):**

> Every engineer on the team should write documentation promptly.

Did "more important than ever" carry a real claim that the reader needs? If the context supports it, restore it briefly: "Now more than ever, every engineer should document promptly."

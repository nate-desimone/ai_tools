# LLM Tells: Making Text Sound Human-Written

Text written or edited by a large language model has recognizable habits. Any one of them is harmless; human writers use an occasional em dash or the word "robust" too. **Clusters** give the text away: a paragraph with a "pinned" key that "lives" somewhere, a trailing ", ensuring ...", and a "The result?" setup reads as machine-written even when every fact is right.

Three rules for fixing them:

1. **Replace generic with specific, not one synonym with another.** Swapping "delve" for "explore" or "robust" for "resilient" keeps the machine voice. Use the field's precise term or the concrete fact.
2. **Use the field's own vocabulary.** Check the document's glossary, the specs it cites, and the codebase. For example, in UEFI code signing, a certificate is *enrolled* in `db`, a measurement is *extended* into a PCR, and data is *sealed* to PCR values.
3. **Aim for normal density, not zero.** Don't strip every em dash or every list of three. Over-correcting creates its own stilted voice.
4. **Unpack before you cut.** Lanham's Official Style pads; LLM text often does the opposite and packs several conditions into one dense phrase (section 7). Fixing that adds words, and a negative Lard Factor is the right result.

`prose_audit.py scan` flags the common tells under "LLM tells" and counts em dashes, arrows, and emoji. It deliberately skips terms of art such as GPIO pins, attack surface, trust anchor, and clock gating.

---

## 1. Things That Come Alive

LLMs describe code, data, and keys as living things or objects that move around. Engineers usually say what the system *does* or where the thing *is*.

| Tell | LLM version | Human version |
|------|-------------|---------------|
| lives in / on / at | The MOR request lives in the MemoryOverwriteRequestControl variable until the next boot. | The MOR request persists in the MemoryOverwriteRequestControl variable until the next boot. |
| sits | The check sits in front of the dispatcher. | The check runs before the dispatcher. |
| owns | SMM owns the private key. | Only SMM can access the private key. |
| wants / expects / knows / sees / cares about | The parser wants a null-terminated string. The OS never sees the key. The loader doesn't care about alignment. | The parser requires a null-terminated string. The OS never has access to the key. The loader ignores alignment. |
| talks to | The driver talks to the BMC. | The driver sends commands to the BMC. / communicates with |
| travels / flows | The hash travels with the capsule. | The capsule includes the hash. |
| hands off | PEI hands off to DXE. | PEI passes control to DXE. (Keep "handoff" and HOB, the PI spec's Hand-Off Block, where they are the spec's terms.) |
| wakes up | The handler wakes up on each SMI. | The handler runs on each SMI. |
| surfaces (verb) | The tool surfaces errors to the caller. | The tool returns errors to the caller. / reports / logs |
| lands | The fix landed last week. | The fix was merged last week. |
| leans on | The design leans on the TPM. | The design relies on the TPM. |
| journey | the boot journey | the boot flow / boot sequence |

"Stays" and "remains" are fine; humans use them for state that persists.

## 2. Vivid Stand-Ins for Precise Terms

LLMs reach for a vivid general verb where the field has a precise one. In technical documents, the precise term is both clearer and more credible.

### "Pin", "pinned", "pinning"

LLMs use "pinned" for almost any kind of trust or binding. Pick the word for the actual mechanism:

| Situation | Precise term |
|-----------|--------------|
| A certificate or key the firmware accepts | trusted, authorized, enrolled (in `db`, `PK`, `KEK`), provisioned |
| A key or hash compiled into the image | embedded, hard-coded, built into |
| A hash on an allow list | stored, recorded, allow-listed |
| Data bound to TPM state | sealed to PCR *n* |
| A measurement | extended into PCR *n*, measured |
| A configuration value that can't change | locked, fixed, write-protected, set at manufacturing |
| A setting tied to a platform | bound to, specific to |

Keep "pin" where it is the established term: **certificate pinning** and **public key pinning** (TLS), **pinned memory** (page-locked DMA buffers), **pinned dependency versions**, **CPU pinning** (affinity), and the noun **pin** (GPIO pin, pin mux).

### Other stand-ins

| Tell | Precise alternatives | Keep when |
|------|---------------------|-----------|
| anchored to / in | rooted in, chained to, verified against | "trust anchor" (RFC 5914) |
| gated on / behind | requires, is enabled only when, depends on | clock gating, power gating |
| wired up / into, plumbed through | connected, passed through, added to | describing physical wiring |
| baked in | built in, compiled in | |
| story ("the security story") | approach, design, support, plan | user stories |
| shape ("the shape of the data") | structure, format, layout | |
| surface area | scope, size, number of interfaces | attack surface |
| load-bearing | essential, required, critical | |
| footgun | pitfall, easy mistake | |
| belt and suspenders | redundant check, defense in depth | |
| escape hatch | override, workaround, fallback | |
| knobs | settings, options | "BIOS knobs" if that's your team's term |
| first-class (support) | full, native, built-in | |
| heavy lifting | most of the work | |
| under the hood | internally | |
| sharp edges | limitations, pitfalls | |
| invariant | rule, guarantee, condition | formal verification, math |
| canonical | standard, official | canonical form, canonical path |
| orthogonal | independent, unrelated | |
| crisp | clear, precise | |
| "the X dance" | the X steps, the X process | |

## 3. LLM-Favored Vocabulary

These words show up far more often in LLM text than in human text. Use the plain word, or better, say the specific thing.

| Tell | Plain alternative |
|------|-------------------|
| delve into | examine, look at, cover |
| tapestry, interplay | mix, interaction |
| a testament to | shows, proves |
| intricate, multifaceted | complex, detailed (or list the parts) |
| realm, landscape, world (of X) | field, area (or cut) |
| navigate (the complexities) | handle, deal with, work through |
| seamless(ly) | (say what: "without a reboot", "with no user action") |
| streamline | simplify, speed up |
| underscore, highlight | show, emphasize |
| pivotal, paramount, crucial, vital | important (or say why it matters) |
| foster, bolster | encourage, strengthen, support |
| harness, leverage, utilize | use |
| unlock (the potential), unleash, empower | enable, allow, let |
| elevate | improve, raise (keep "elevated privileges") |
| embark on | start, begin |
| showcase | show |
| meticulous(ly) | careful(ly), thorough(ly) |
| holistic | complete, end-to-end, system-wide |
| robust | (say what it survives: "recovers from power loss") |
| comprehensive | complete, full (or cut) |
| groundbreaking, transformative, game-changing | (cut; give the result) |
| myriad, plethora, a wide range of | many, various (or list them) |
| beacon, cornerstone, linchpin, bedrock, hallmark | (say what it does) |
| nuanced | subtle (or explain the nuance) |
| genuinely, truly, honestly | (cut) |
| notably, importantly, crucially, interestingly | (cut) |
| additionally, moreover, furthermore | also, and (or cut) |
| ensure | make sure, or name the mechanism: checks, verifies, guarantees |
| key (insight, takeaway, consideration, aspect) | (say what it is) |
| plays a crucial / pivotal / key role in | (say what it does) |
| quietly | silently (or say what happens) |
| cleanly, neatly, squarely, materially, meaningfully | (cut or be specific) |
| inherently, fundamentally | (usually cut) |

## 4. Sentence Patterns

| Pattern | Example | Fix |
|---------|---------|-----|
| "Not X, but Y" framing | This isn't just a fix; it's a new approach. / X isn't about Y. It's about Z. | State Y. Keep the contrast only if readers really hold belief X. |
| Trailing benefit clause | ...signed by the vendor, ensuring a robust chain of trust. | Cut it, or make it a separate claim with evidence. |
| Vague "This" | This ensures that... / This means... | Name the subject: "The lock bit ensures..." |
| Rhetorical setup | The result? Faster boots. / Here's the catch: ... / Why does this matter? | State it: "Boot time dropped 18%." |
| Reflexive triplets | fast, secure, and reliable | Use the number of items you have, often one or two. |
| Paired synonyms | clear and concise; robust and scalable | Pick one. |
| Even-handed hedging | While X has benefits, it also has drawbacks. | Say which benefits and drawbacks matter, and recommend. |
| Sweeping opener | In today's... / In a world where... / Whether you're X or Y... / From X to Y... | Cut; start with the point. |
| Signpost opener | Let's dive in. / Let's break it down. / Now, let's look at... | Cut; just start. |
| Summary closer | In summary, ... / Overall, ... / Ultimately, ... / In essence, ... | Cut the restatement; end on the last real point. |
| Pithy closer | ...And that makes all the difference. | Cut. |
| Chatbot leakage | Great question! / Certainly! / I hope this helps. / Let me know if you have any questions. | Cut. (In an email, keep a real, specific ask.) |
| Unattributed authority | Experts agree... / Studies show... / It is widely recognized... | Cite the source or cut. |
| Signpost colon | The fix: ... / The problem: ... / Bottom line: ... | Write a sentence: "The fix moves the check before the copy." |
| No contractions anywhere | does not, was not, is not in every sentence | Match the author: doesn't, wasn't, isn't where the author uses them. Never contract normative keywords (shall not, must not). |

## 5. Punctuation and Formatting

- **Em dashes (—).** LLMs use several per paragraph. Allow about one per paragraph. Replace the rest with a comma, colon, parentheses, or a period, without creating run-on sentences.
- **Unicode arrows, check marks, and emoji** (→, ⇒, ✓, ✅, 🚀) in prose, headings, or bullets. Use words ("becomes", "then", "yes") unless the document's house style uses symbols. Comparison tables are fine.
- **Bold everywhere.** Bold lead-ins on every bullet ("**Speed:** ...") and bold phrases scattered through prose. Save bold for defined terms and real warnings.
- **Over-structuring.** Headings over two-paragraph sections, bullets for ideas that connect, or a "Key Takeaways" or "Conclusion" section in a short document. Use prose when ideas build on each other; use bullets for parallel, separable items.
- **Heading case.** LLMs title-case every heading. Follow the document's existing convention.
- **Parentheticals.** An "(e.g., ...)" or aside in every sentence. Keep the ones the reader needs.

## 6. Content Tells

These are the hardest to fix and matter most.

- **Generic claims where the author would give specifics.** "Significantly improves performance" vs. "cuts S3 resume from 410 ms to 260 ms". Use the specifics the document supports. **Never invent them**; flag them for the author instead.
- **Uniform confidence.** Every claim is equally certain or equally hedged. Humans are confident about what they tested and hedge what they didn't.
- **No one does anything.** Actions happen without actors. Name them (Paramedic Method step 3).
- **Uniform rhythm.** Every sentence runs 15–25 words and every paragraph has three or four sentences. Vary them (Paramedic Method step 8).
- **Over-explaining.** The text defines basics the audience knows, or restates the question before answering it. Write for the actual reader.
- **Balanced to a fault.** The text lists pros and cons but never recommends. If the document should take a position, take it.

## 7. Over-Compression

LLMs pack conditions, states, and actors into dense noun phrases. The text looks tight, but the reader has to unpack the logic and may unpack it wrong. This matters most in specifications, where each condition must be testable.

| Tell | LLM version | Human version |
|------|-------------|---------------|
| Condition dressed as a noun + "in which" | updated only in a boot in which BDS reports... | updated only if BDS reports... |
| Same, other forms | in cases where... / on any platform on which... / in a scenario where... | if... / when... / on platforms that... |
| Modifiers on the wrong noun | BDS reports a deployed, secure boot | BDS reports that the platform is in Deployed Mode and Secure Boot is enabled |
| An event as the actor | A boot in which the UEFI Shell loaded updates nothing. | SMM does not update db. |
| Ambiguous intransitive verb | the UEFI Shell loaded | the UEFI Shell was loaded / the Boot Manager loaded the UEFI Shell |
| Vague negative object | updates nothing / changes nothing | does not update db |
| Implicit scope | a boot in which... / afterwards | during the current boot cycle / after `ExitBootServices` |
| Stand-in for a defined term | ...updates it / ...updates the database | ...updates db |
| Restriction stacked on "no" | only if no UEFI Shell was loaded | only if the UEFI Shell wasn't loaded |
| Negative rule split from its requirement | Release K only if A and B. If C happened, K is not released. | Release K only if A, B, and C didn't happen. |

A boot is not deployed; the platform is. A boot does not update db; SMM does. When an adjective or verb attaches to something that can't have that property or do that action, find the real subject.

The passive is clearer than an ambiguous intransitive verb. "The application loaded" could mean it finished loading or that it loaded something else. "The application was loaded" has one reading.

**How to unpack a compressed sentence:**

1. List every condition and every action the sentence contains.
2. Give each condition its own subject and verb: "the platform is in Deployed Mode", "Secure Boot is enabled".
3. Write conditions with if, only if, when, or unless. Put "then" before the consequence when the condition is long.
4. Name the component that acts at least once (SMM, the Boot Manager, the UEFI OS Loader, etc.).
5. Fold separate "does not" or "shall not" rules about the same action into the "only if" conditions, and negate their verbs ("the UEFI Shell wasn't loaded").
6. With three or more conditions, prefer a list: "only if:" followed by the items, with "; and" or "; or" before the last one so the reader sees how they combine.
   - Use one sentence when a list doesn't fit, such as inside an existing list item (no nested lists), in a table cell, or in a format that avoids lists.
   - In the sentence, put a condition introduced by a reporting verb ("BDS reports that...") last so the verb can't seem to cover the others.
7. If unpacking reveals an ambiguity, such as which component checks which condition, ask the author. Don't guess.

---

## Worked Examples

### Example 1: Anthropomorphic metaphor

**Before:** When the OS sets the MOR bit, the request lives in the MemoryOverwriteRequestControl variable until the next boot, when firmware clears memory before any option ROM or OS loader runs.

**After:** When the OS sets the MOR bit, the request persists in the MemoryOverwriteRequestControl variable until the next boot, when firmware clears memory before any option ROM or OS loader runs.

### Example 2: "Pinned" plus a trailing benefit clause (27 → 18 words, LF 33%)

**Before:** The update certificate is pinned in the firmware volume, so the platform only accepts capsules signed by the pinned key — ensuring a robust, seamless chain of trust.

**After:** The firmware volume contains the authorized update certificate, and the platform accepts only capsules signed with its key.

The trailing clause added no information, so it goes. If the chain of trust matters to the reader, describe it in its own sentence.

### Example 3: Framing, rhetorical setup, and a triplet (26 → 14 words, LF 46%)

**Before:** This isn't just a performance fix — it's a fundamental shift in how we handle SMIs. The result? Faster boots, fewer timeouts, and a more robust platform.

**After:** The fix changes how we handle SMIs, which shortens boot time and reduces timeouts.

**Please review:** Add the measured boot-time and timeout numbers. The revision should state them, but the original doesn't give them. The revision also drops "a more robust platform" because the claim names nothing testable. Restore it if you can say what the platform now survives.

### Example 4: Over-compressed requirement (52 → 61 words, LF −17%)

**Before:**

> db shall be updated only in a boot in which BDS reports a deployed, secure boot, and every image loaded before ExitBootServices came from a firmware volume or was signed with a certificate enrolled in db. A boot in which the UEFI Shell loaded updates nothing (section 4.2).

**After:**

> db shall be updated only if BDS reports the platform is in Deployed Mode, Secure Boot is enabled, and every image loaded before ExitBootServices came from a firmware volume or was signed with a certificate enrolled in db. If the UEFI Shell was loaded during the current boot cycle, then SMM does not update db (section 4.2).

The revision is longer and better. It:

- turns "in a boot in which" into "only if";
- splits "a deployed, secure boot" into two conditions with real subjects;
- replaces the boot as the actor with SMM;
- changes "loaded" to "was loaded";
- names db instead of "nothing";
- makes the scope explicit with "the current boot cycle".

**Tighter: one requirement (54 words; LF −4% vs. the original, 12% vs. the first revision).** Folding the UEFI Shell rule into the conditions removes a sentence and settles which conditions BDS reports.

List form, the usual choice:

> SMM shall update db only if:
>
> - the UEFI Shell wasn't loaded during this boot (section 4.2);
> - every image loaded before ExitBootServices came from a firmware volume or was signed with a certificate enrolled in db; and
> - BDS reports that the platform is in Deployed Mode and Secure Boot is enabled.

Prose form, for when the requirement already sits in a list item, a table cell, or a format without nested lists:

> SMM shall update db only if the UEFI Shell wasn't loaded during this boot (section 4.2), every image loaded before ExitBootServices came from a firmware volume or was signed with a certificate enrolled in db, and BDS reports that the platform is in Deployed Mode and Secure Boot is enabled.

Why these read as human-written:

- **The negation is on the verb:** "the UEFI Shell wasn't loaded". The machine-sounding "only if no UEFI Shell was loaded" stacks "only if" on "no".
- **BDS comes last,** so "reports that" can't seem to cover the image check.
- **The section reference sits right after** the condition it supports.


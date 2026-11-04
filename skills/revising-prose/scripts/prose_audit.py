#!/usr/bin/env python3
"""Prose audit for the revising-prose skill.

scan:    Paramedic Method markup and wordiness metrics for one document.
compare: Lard Factor and before/after metrics for an original and a revision.

Markdown aware: front matter, fenced code, inline code, URLs, images, HTML
tags, and HTML comments are excluded so word counts reflect prose only.
"""

from __future__ import annotations

import argparse
import re
import statistics
import sys
import textwrap
from dataclasses import dataclass, field
from pathlib import Path

WORD_RE = re.compile(r"[A-Za-z0-9]+(?:['\-_][A-Za-z0-9]+)*")

PREPOSITIONS = frozenset(
    """
    about above across after against along amid among amongst around at before behind
    below beneath beside besides between beyond by concerning despite during except for
    from in inside into near of onto outside over per regarding through throughout to
    toward towards under underneath until upon via with within without
    """.split()
)

BE_FORMS = frozenset(
    """
    am is are was were be been being isn't aren't wasn't weren't it's that's there's
    here's what's who's where's he's she's i'm you're we're they're
    """.split()
)

IRREGULAR_PARTICIPLES = frozenset(
    """
    begun bent bound bought broken brought built burnt caught chosen cut dealt done drawn
    driven eaten fallen felt fought found forgotten forgiven frozen given gone grown held
    hidden hit hung hurt kept known laid led left lent let lost made meant met paid put
    quit read rewritten run said seen sent set shaken shown shut sold sought spent split
    spoken spread stolen struck stuck taken taught thrown told thought torn understood
    undertaken withdrawn won worn written
    """.split()
)

NOMINALIZATION_SUFFIX_RE = re.compile(r"(?:tion|sion|ment|ance|ence|ancy|ency|ity|ness)s?$")
NOMINALIZATION_EXCLUDE = frozenset(
    """
    audience audiences business businesses comment comments community department
    departments document documents element elements entity entities environment
    environments evidence experience experiences government instance instances moment
    moments question questions science section sections security segment segments sentence
    sentences quality university
    """.split()
)
NOUN_OF_RE = re.compile(
    r"\b(?:the|a|an)\s+([a-z]+(?:tion|sion|ment|ance|ence|ancy|ency|ity|ness)s?)\s+of\b", re.I
)

EXPLETIVE_RES = (
    re.compile(
        r"^(?:there|here)(?:'s|\s+(?:is|are|was|were|will\s+be|would\s+be|has\s+been|"
        r"have\s+been|had\s+been|seems?\s+to\s+be|appears?\s+to\s+be))\b",
        re.I,
    ),
    re.compile(
        r"^it(?:'s|\s+(?:is|was|will\s+be|would\s+be|has\s+been|had\s+been|seems|seemed|"
        r"appears|appeared))\s+(?:\S+\s+){0,3}?(?:that|to|whether|how|why|when)\b",
        re.I,
    ),
)

NEGATIVE_TOKENS = frozenset("not no never none nothing nobody neither nor nowhere cannot".split())
NEGATED_WORD_RE = re.compile(
    r"\b(?:not|never|hardly|scarcely)\s+(?:(?:a|an|so|entirely|wholly|completely|totally|"
    r"altogether)\s+)?([a-z]+)\b",
    re.I,
)
NEGATIVE_PREFIX_WORDS = frozenset(
    """
    impossible improbable implausible improper impractical immaterial imperfect immature
    inaccurate inadequate inappropriate incapable incomplete inconsistent inconsiderable
    inconceivable incorrect ineffective inefficient inelegant inevitable inexpensive
    infrequent infrequently insignificant insubstantial insufficient invalid illegal
    illegitimate illogical irrational irregular irrelevant irresponsible irreversible
    disagree disagreeable dishonest dislike disliked dissimilar displeased disapprove
    disapproved
    """.split()
)
UN_NOT_NEGATIVE = ("until", "unless", "under", "uni", "unto")
NEG_CONSTRUCTION_RE = re.compile(
    r"\b(?:not|never|cannot|can't|won't|don't|doesn't|didn't|isn't|aren't|wasn't|weren't|no)"
    r"\b(?:\s+\S+){0,6}?\s+(?:without|unless)\b"
    r"|\bfail(?:s|ed|ing)?\s+to\s+not\b"
    r"|\b(?:not|no)\s+(?:\S+\s+){0,3}?not\b",
    re.I,
)

# (pattern, suggestion). Specific patterns come before general ones so they win overlaps.
LARD_PHRASES = [
    (r"in order to|so as to|in an effort to|with the goal of", "to"),
    (r"in order for", "for"),
    (r"(?:due|owing) to the fact that|in light of the fact that|for the reason that", "because"),
    (r"in light of", "given / because of"),
    (r"(?:with)?in the context of", "in"),
    (r"(?:despite|in spite of) the fact that", "although"),
    (r"the fact that", "that (or rewrite)"),
    (r"at this point in time|at the present time", "now"),
    (r"at that point in time", "then"),
    (r"in the event that", "if"),
    (r"in the near future", "soon"),
    (r"prior to", "before"),
    (r"subsequent to", "after"),
    (r"until such time as", "until"),
    (r"for the purpose of", "to / for"),
    (r"with (?:regard|respect) to|in (?:regard|reference|relation) to", "about / on / for"),
    (r"in terms of", "rewrite (often 'in' or 'for')"),
    (r"on the basis of", "by / from / because of"),
    (r"by means of|by virtue of", "by"),
    (r"in accordance with", "under / by / following"),
    (r"in conjunction with", "with"),
    (r"with the exception of", "except"),
    (r"in the absence of", "without"),
    (r"in excess of", "more than"),
    (r"(?:in )?close proximity(?: to)?", "near"),
    (r"on the part of", "by"),
    (r"(?:the question )?as to whether", "whether"),
    (r"the (?:way|manner) in which", "how"),
    (r"whether or not", "whether (keep 'or not' only if it means 'regardless')"),
    (r"at all times", "always"),
    (r"a (?:large |great |significant )?number of", "many / several / some"),
    (r"(?:the|a) majority of", "most"),
    (r"a great deal of", "much"),
    (r"a total of", "(cut)"),
    (r"(?:is|are|was|were|be|been) able to", "can / could"),
    (r"ha(?:s|ve|d) the (?:ability|capability|capacity) to", "can / could"),
    (r"ha(?:s|ve|d) the potential to|it is possible that", "can / may / might"),
    (r"it is (?:necessary|essential|imperative) that", "must"),
    (r"(?:is|are|was|were) in need of", "need(s)"),
    (r"(?:is|are|was|were) dependent (?:on|upon)", "depend(s) on"),
    (r"(?:is|are|was|were) indicative of", "indicate(s)"),
    (r"(?:is|are|was|were) (?:reflective of|a reflection of)", "reflect(s)"),
    (r"(?:is|are|was|were) supportive of", "support(s)"),
    (r"(?:is|are|was|were) in agreement(?: with)?", "agree(s) (with)"),
    (r"(?:is|are|was|were) in compliance with", "comply / complies with"),
    (r"(?:is|are|was|were) applicable to", "apply / applies to"),
    (r"(?:is|are|was|were) in possession of", "have / has"),
    (r"(?:is|are|was|were) of the opinion that", "think(s) / believe(s)"),
    (r"(?:is|are|was|were|be|been|become|becomes|became) (?:an? |the )?(?:active |heavy |major )?"
     r"(?:participants? in|users? of|contributors? to|providers? of|supporters? of|believers? in)",
     "use the verb (participate, use, contribute, provide, support, believe)"),
    (r"was in attendance|were in attendance", "attended"),
    (r"serv(?:e|es|ed|ing) to", "(cut; use the main verb)"),
    (r"(?:make|makes|made|making) (?:a |the )?decisions?", "decide(d)"),
    (r"(?:make|makes|made|making) (?:a |the )?recommendations?", "recommend(ed)"),
    (r"(?:make|makes|made|making) (?:an )?improvements? (?:to|in)", "improve(d)"),
    (r"(?:give|gives|gave|given|giving) consideration to|take into consideration|took into consideration", "consider"),
    (r"(?:conduct|conducts|conducted|perform|performs|performed) an? (?:analysis|investigation|review|evaluation|assessment) of",
     "analyze / investigate / review / evaluate / assess"),
    (r"(?:provide|provides|provided|providing) (?:assistance|help|support) (?:to|for|with)", "help / support"),
    (r"(?:make|makes|made|making) use of", "use"),
    (r"utiliz(?:e|es|ed|ing|ation)", "use"),
    (r"leverag(?:e|es|ed|ing)", "use"),
    (r"(?:come|comes|came|coming) to (?:a|the) conclusion", "conclude(d)"),
    (r"(?:reach|reaches|reached|reaching) (?:an )?agreement", "agree(d)"),
    (r"(?:has|have|had) an? (?:impact|effect) on", "affect(s)"),
    (r"(?:achieve|achieves|achieved|achieving) an? (?:reduction|increase|improvement) in", "reduce / increase / improve"),
    (r"it is (?:important|worth|interesting) (?:to note|noting|to mention|mentioning) that", "(cut)"),
    (r"it should be (?:noted|mentioned|emphasized|pointed out) that", "(cut)"),
    (r"needless to say|it goes without saying(?: that)?", "(cut)"),
    (r"the fact of the matter is(?: that)?", "(cut)"),
    (r"the point (?:I|we) (?:want|wish|would like|am trying|are trying) to make is(?: that)?", "(cut)"),
    (r"it is (?:widely|well|generally|commonly|universally) (?:known|accepted|understood|believed|recognized|agreed)(?: that)?",
     "(cut, or say who knows/accepts it)"),
    (r"what (?:this|that) means is(?: that)?", "(cut / rewrite)"),
    (r"the reason (?:why )?(?:\S+ ){0,6}?(?:is|was) (?:because|that)", "because"),
    (r"as (?:previously|already) (?:mentioned|stated|noted|discussed)|as mentioned (?:above|earlier|previously)", "(cut)"),
    (r"the purpose of this (?:document|section|paper|report|email|memo) is to", "(cut; just do it)"),
    (r"in this (?:document|section|paper|report),? (?:we|I) will", "(cut; just do it)"),
    (r"I would like to take this opportunity to|I would like to", "(cut) / I"),
    (r"in today's (?:\S+ ){0,3}?(?:world|landscape|environment|age|era|market)", "(cut)"),
    (r"for all intents and purposes", "(cut) / in effect"),
    (r"at the end of the day", "(cut) / ultimately"),
    (r"when it comes to", "for / in"),
    (r"please do not hesitate to|please don't hesitate to", "please"),
    (r"first of all", "first"),
    (r"last but not least", "finally"),
    (r"in the (?:process|course) of", "while / during"),
    (r"(?:on|in) an? (?:daily|weekly|monthly|yearly|annual|regular|consistent|ongoing) basis",
     "daily / weekly / monthly / yearly / regularly / consistently"),
    (r"in a timely (?:manner|fashion)", "promptly"),
    (r"in an? \w+ (?:manner|fashion)", "use the adverb (e.g., 'quickly')"),
    (r"(?:of|in) an? (?:\w+ )?nature", "use the adjective"),
    (r"(?:large|small|big) in size|\w+ in (?:color|colour|shape)", "drop the category"),
    (r"in the (?:area|field|realm|domain) of", "in"),
    (r"(?:period|length) of time", "period / time"),
    (r"each and every", "each / every"),
    (r"first and foremost", "first"),
    (r"any and all", "any / all"),
    (r"full and complete|true and accurate|various and sundry|null and void|basic and fundamental", "pick one word"),
    (r"past (?:history|experience)", "history / experience"),
    (r"end result|final (?:outcome|result)", "result / outcome"),
    (r"future plans|plan ahead|planned ahead|advance planning", "plan(s) / planning"),
    (r"basic (?:fundamentals|essentials)", "fundamentals / essentials"),
    (r"completely (?:eliminate|eliminated|eliminates|finish|finished|destroy|destroyed)", "drop 'completely'"),
    (r"(?:very|totally|completely|quite|most) unique", "unique"),
    (r"(?:joint|mutual) (?:collaboration|cooperation)", "collaboration / cooperation"),
    (r"new innovations?", "innovation(s)"),
    (r"general consensus|consensus of opinion", "consensus"),
    (r"still remains?|continues? to remain", "remain(s)"),
    (r"(?:revert|return)(?:s|ed)? back", "revert / return"),
    (r"(?:combine|merge|join|connect)(?:s|d|ed)? together", "combine / merge / join / connect"),
    (r"(?:repeat|reiterate)(?:s|d|ed)? again", "repeat / reiterate"),
    (r"(?:currently|presently) ongoing", "ongoing"),
    (r"(?:true|actual) facts?", "fact(s)"),
    (r"absolutely (?:essential|necessary|critical|vital)", "essential / necessary / critical / vital"),
    (r"unexpected surprise|added bonus|free gift|sum total|each individual", "drop the redundant word"),
    (r"ATM machine|PIN number|LCD display", "ATM / PIN / LCD"),
    (r"going forward|moving forward", "(cut) / from now on"),
    (r"circle back|touch base", "follow up / talk"),
    (r"drive alignment|achieve alignment", "agree / align"),
    (r"(?:is|was|are|were|seems?|feels?|looks?) (?:kind|sort) of|kinda|sorta", "hedge: cut, or be precise"),
    (r"I think|I believe|I feel|in my opinion|it seems that", "cut unless contrasting opinion with fact"),
    (r"very|really|quite|rather(?!\s+than)|extremely|truly|totally|incredibly|super", "intensifier: cut or use a stronger word"),
    (r"basically|essentially|actually|literally|obviously|of course|somewhat", "filler: cut if it adds nothing"),
]
LARD_RES = [(re.compile(rf"\b(?:{p})\b", re.I), s) for p, s in LARD_PHRASES]


def _words(pattern: str) -> str:
    return rf"\b(?:{pattern})\b"


# Fixed phrases and spec limits where "no" is idiomatic.
_NO_IDIOMS = (r"(?!-)(?! (?:longer|one|matter|more|less|fewer|later|earlier|sooner|greater|smaller|larger|bigger)\b)"
              r"(?=\s+\w)")


# Habits that make text sound machine-written. Patterns avoid terms of art where possible
# (GPIO pins, attack surface, trust anchor, clock gating, test harness, elevated privileges).
LLM_TELLS = [
    (_words(r"pinned|pinning|pins? (?:it|them|down)"),
     "trusted / authorized / enrolled / provisioned / stored (keep certificate pinning, pinned memory, pinned versions)"),
    (_words(r"anchor(?:s|ed|ing)? (?:to|in|on)"), "rooted in / chained to / verified against (keep 'trust anchor')"),
    (_words(r"lives? (?:only )?(?:in|on|at|inside|within|under|behind|outside|alongside)|living (?:in|on|at|inside|within)"),
     "is / is stored in / exists in / is defined in"),
    (_words(r"sit(?:s|ting)? (?:in|on|at|behind|between|above|below|inside|within|alongside|underneath|on top of|in front of)"),
     "is / runs / is placed"),
    (_words(r"surfac(?:es|ed|ing)|surface (?:the|an?|any|this|that|these|those|it|them|errors?|issues?|problems?|failures?|warnings?)"),
     "report / show / log / return"),
    (_words(r"gat(?:e|es|ed|ing) (?:on|behind)"), "requires / is enabled only when"),
    (_words(r"wir(?:e|es|ed|ing) (?:up|into|through)|plumb(?:s|ed|ing)? (?:through|into)"), "connect / pass / add"),
    (_words(r"land(?:s|ed)? (?:in|on)|landed"), "was merged / was released / arrives"),
    (_words(r"lean(?:s|ed|ing)? (?:on|into)"), "rely on / use"),
    (_words(r"bak(?:e|es|ed|ing) (?:in|into)|baked-in"), "build into / built-in"),
    (_words(r"(?:security|upgrade|migration|testing|test|deployment|signing|update|debugging|support|recovery|provisioning) story"
            r"|the \w+ dance"),
     "approach / process / support"),
    (_words(r"load-bearing|footguns?|belt[- ]and[- ]suspenders|escape hatch(?:es)?|sharp edges|heavy lifting|under the hood"
            r"|front and center|north star|table stakes|move the needle|double down|low-hanging fruit|first-class"),
     "cliche: say it plainly"),
    (_words(r"delv(?:e|es|ed|ing)|tapestry|testament|intricate|intricacies|multifaceted|realm|myriad|plethora|interplay"
            r"|beacon|cornerstone|linchpin|bedrock|hallmark|paradigm|synerg(?:y|ies)"),
     "LLM-favored word: use a plain one"),
    (_words(r"seamless(?:ly)?|streamlin(?:e|es|ed|ing)|underscor(?:e|es|ed|ing)|pivotal|paramount|crucial|foster(?:s|ed|ing)?"
            r"|bolster(?:s|ed|ing)?|harness(?:es|ed|ing)|harness (?:the|its|their|this)"
            r"|unlock(?:s|ed|ing)? (?:the )?(?:potential|power|value|new)|unleash(?:es|ed|ing)?|empower(?:s|ed|ing)?"
            r"|embark(?:s|ed|ing)?|showcas(?:e|es|ed|ing)|meticulous(?:ly)?|holistic(?:ally)?|vibrant|groundbreaking"
            r"|transformative|game[- ]chang(?:er|ing)|ever-(?:evolving|changing|growing)|commendable|noteworthy|invaluable"
            r"|nuanced|crisp|orthogonal|robust(?:ly|ness)?|comprehensive"),
     "LLM-favored word: use a plain or specific one"),
    (_words(r"genuinely|honestly|to be honest|quietly|cleanly|neatly|squarely|materially|meaningfully|inherently|fundamentally"),
     "filler adverb: cut, or say what you mean"),
    (r"^(?:crucially|importantly|notably|interestingly|concretely|ultimately|overall|additionally|moreover|furthermore"
     r"|in essence|at its core|simply put|put simply|in short|in summary|in conclusion|to summarize|that said"
     r"|that being said)\b,?",
     "signpost opener: cut"),
    (r"^(?:certainly|absolutely|great question)[!,.]|"
     + _words(r"i hope this helps|let me know if you have any (?:other |further )?questions|feel free to (?:reach out|ask)"
              r"|happy to help"),
     "chatbot phrasing: cut"),
    (_words(r"(?:let's|let us) (?:dive|delve|explore|unpack|break (?:it|this|that) down|take a (?:closer )?look|walk through)"
            r"|dive (?:in|into|deeper)|deep[- ]dive"),
     "cut; just start"),
    (_words(r"here's (?:the thing|why|how|what|the catch|the kicker|the deal)"
            r"|the (?:result|catch|answer|kicker|upshot|bottom line|short answer|good news|bad news|takeaway)(?=\s*[?:])"),
     "rhetorical setup: state it directly"),
    (_words(r"(?:isn't|is not|aren't|are not|wasn't|was not|it's not|this isn't) (?:just|merely|simply|only|about)"),
     "'not X, but Y' framing: state Y directly"),
    (r",\s+(?:ensuring|enabling|allowing|highlighting|underscoring|reflecting|demonstrating|showcasing|fostering|paving"
     r"|reinforcing|emphasizing|making it)\b|\u2014\s*(?:ensuring|enabling|allowing|making)\b",
     "trailing benefit clause: cut, or make it a separate claim with evidence"),
    (r"^this (?:ensures|allows|enables|means|makes|helps|provides|guarantees)\b", "vague 'This': name the thing"),
    (_words(r"plays? an? (?:crucial|pivotal|key|vital|critical|significant|important|central) role(?: in)?"),
     "say what it does"),
    (_words(r"a (?:wide|broad|diverse|rich) (?:range|array|variety|spectrum) of"), "many / various (or list them)"),
    (_words(r"key (?:insights?|takeaways?|considerations?|aspects?|benefits?)"), "say what it is"),
    (_words(r"it's (?:worth|important to) (?:noting|mentioning|note|mention|calling out|flagging|highlighting|pointing out)"
            r"(?: that)?|worth (?:calling out|flagging|highlighting)"),
     "(cut)"),
    (_words(r"shed(?:s|ding)? light on|pav(?:e|es|ed|ing) the way|stands? as an? (?:testament|example|reminder)|a testament to"
            r"|navigat(?:e|es|ed|ing) (?:the )?(?:complexities|landscape|challenges|world)"
            r"|in the (?:ever-\w+ )?(?:landscape|realm|world) of|in a world where|whether you're"),
     "cliche: say it plainly"),
    (_words(r"(?:a|an|any|each|every|the) (?:\w+ ){1,2}?(?:in|on|during|under|for|at) which|in (?:cases|situations|scenarios) where"
            r"|in (?:a|any|the) (?:case|situation|scenario) where"),
     "condition as noun + clause: use if / only if / when"),
    (_words(r"[a-z]+(?:s|ed) nothing"), "vague negative: 'does not <verb> <named object>'"),
    (_words(r"only if no|unless no|except (?:when|if) no"), "negate the verb: 'only if a <noun> wasn't <verb>'"),
    (_words(r"with no" + _NO_IDIOMS), "use 'without': 'with no X' -> 'without X' / 'without any X'"),
    (_words(r"(?:[a-z]+(?:s|ed)|had|held|made|took|kept|left|gave|saw|got|found|sent|have|hold|contain|carry|require|need"
            r"|take|make|provide|offer|perform|expose|store|retain|keep|leave|see|get|receive|return|produce|include|use"
            r"|allow|give|show|report|incur|add|introduce|impose|be) no" + _NO_IDIOMS),
     "negated noun: negate the verb ('doesn't hold any X', 'shall not contain X')"),
    (_words(r"at no (?:point|time)|in no way|under no circumstances|by no means"), "never / not"),
]
LLM_RES = [(re.compile(p, re.I), s) for p, s in LLM_TELLS]
NORMATIVE_RE = re.compile(r"\b(?:shall|must|required)\b", re.I)
DASH_RE = re.compile(r"\u2014|\s--\s|\s\u2013\s")
SYMBOL_RE = re.compile("[\u2190-\u21FF\u2600-\u27BF\u2B00-\u2BFF\U0001F000-\U0001FAFF]")

ABBREVIATIONS = frozenset(
    "e.g i.e etc vs mr mrs ms dr prof inc ltd jr sr st fig approx dept al cf u.s".split()
)
BOUNDARY_RE = re.compile(r"[.!?]+[\"')\]]*\s+")
SENTENCE_START_CHARS = "\"'([*_"

QUOTE_TRANSLATION = str.maketrans({"\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"'})
HTML_COMMENT_RE = re.compile(r"<!--.*?-->", re.S)
FENCE_RE = re.compile(r"^\s*(`{3,}|~{3,})")
HEADING_RE = re.compile(r"^\s{0,3}#{1,6}\s+(.*?)\s*#*\s*$")
HR_RE = re.compile(r"^\s{0,3}([-*_=])(?:\s*\1){2,}\s*$")
LIST_RE = re.compile(r"^\s*(?:[-*+]|\d+[.)])\s+(?:\[[ xX]\]\s+)?")
TABLE_SEP_RE = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(?:\|\s*:?-{2,}:?\s*)*\|?\s*$")
QUOTE_RE = re.compile(r"^\s*(?:>\s?)+")
INLINE_CLEANERS = (
    (re.compile(r"`[^`\n]*`"), ""),
    (re.compile(r"!\[[^\]\n]*\]\([^)\n]*\)"), ""),
    (re.compile(r"\[([^\]\n]*)\]\([^)\n]*\)"), r"\1"),
    (re.compile(r"\[([^\]\n]*)\]\[[^\]\n]*\]"), r"\1"),
    (re.compile(r"<https?://[^>\s]+>"), ""),
    (re.compile(r"https?://\S+"), ""),
    (re.compile(r"</?[A-Za-z][^>\n]*>"), ""),
    (re.compile(r"\*{1,3}|(?<!\w)_{1,3}|_{1,3}(?!\w)|~~"), ""),
)

LONG_SENTENCE = 30
MONOTONE_SPREAD = 6
MONOTONE_MIN_SENTENCES = 4


@dataclass
class Segment:
    line: int
    text: str
    kind: str  # "prose", "heading", or "cell"


@dataclass
class Sentence:
    line: int
    kind: str
    text: str
    words: list[str]
    preps: list[int] = field(default_factory=list)
    be: list[int] = field(default_factory=list)
    noms: list[int] = field(default_factory=list)
    passive: int = 0
    noun_of: list[str] = field(default_factory=list)
    chains: list[list[str]] = field(default_factory=list)
    lard: list[tuple[str, str]] = field(default_factory=list)
    expletive: bool = False
    negatives: int = 0
    double_neg: list[str] = field(default_factory=list)
    llm: list[tuple[str, str]] = field(default_factory=list)
    dashes: int = 0
    symbols: int = 0

    @property
    def score(self) -> float:
        return (
            len(self.preps)
            + 1.5 * len(self.be)
            + len(self.noms)
            + 2 * self.passive
            + 2 * len(self.noun_of)
            + 3 * len(self.chains)
            + 2 * len(self.lard)
            + 3 * self.expletive
            + 3 * len(self.double_neg)
            + 2 * len(self.llm)
            + self.dashes
            + self.symbols
            + max(0, len(self.words) - 25) / 5
        )


@dataclass
class Report:
    path: Path
    sentences: list[Sentence]
    blocks: list[list[Sentence]]


def is_nominalization(word: str) -> bool:
    return len(word) >= 7 and word not in NOMINALIZATION_EXCLUDE and bool(NOMINALIZATION_SUFFIX_RE.search(word))


def is_participle(word: str) -> bool:
    return (len(word) > 3 and word.endswith("ed")) or word in IRREGULAR_PARTICIPLES


def is_negative_word(word: str) -> bool:
    w = word.lower()
    if w.startswith("non") and len(w) > 5:
        return True
    if w.startswith("un") and len(w) > 4 and not w.startswith(UN_NOT_NEGATIVE):
        return True
    return w in NEGATIVE_PREFIX_WORDS


def find_chains(lower: list[str], preps: list[int], max_gap: int = 4, min_len: int = 3) -> list[list[str]]:
    chains, run = [], []
    for idx in preps:
        if run and idx - run[-1] - 1 > max_gap:
            if len(run) >= min_len:
                chains.append([lower[k] for k in run])
            run = []
        run.append(idx)
    if len(run) >= min_len:
        chains.append([lower[k] for k in run])
    return chains


def find_phrases(text: str, regexes: list[tuple[re.Pattern, str]]) -> list[tuple[str, str]]:
    hits = []
    for regex, suggestion in regexes:
        for m in regex.finditer(text):
            hits.append((m.start(), m.end(), m.group(0), suggestion))
    hits.sort(key=lambda h: (h[0], -(h[1] - h[0])))
    kept: list[tuple[int, int, str, str]] = []
    for h in hits:
        if not any(h[0] >= k[0] and h[1] <= k[1] for k in kept):
            kept.append(h)
    return [(h[2], h[3]) for h in kept]


def find_double_negatives(text: str) -> list[str]:
    hits = [m.group(0) for m in NEGATED_WORD_RE.finditer(text) if is_negative_word(m.group(1))]
    hits.extend(m.group(0) for m in NEG_CONSTRUCTION_RE.finditer(text))
    return hits


def analyze_sentence(line: int, kind: str, text: str) -> Sentence:
    words = WORD_RE.findall(text)
    lower = [w.lower() for w in words]
    s = Sentence(line, kind, text, words)
    for i, w in enumerate(lower):
        if w in PREPOSITIONS:
            s.preps.append(i)
        elif w in BE_FORMS:
            s.be.append(i)
        elif is_nominalization(w):
            s.noms.append(i)
    for i in s.be:
        j = i + 1
        if j < len(lower) and lower[j].endswith("ly"):
            j += 1
        if j < len(lower) and is_participle(lower[j]):
            s.passive += 1
    s.chains = find_chains(lower, s.preps)
    s.noun_of = [m.group(1) for m in NOUN_OF_RE.finditer(text)]
    s.lard = find_phrases(text, LARD_RES)
    s.expletive = any(r.match(text.lstrip(SENTENCE_START_CHARS + " ")) for r in EXPLETIVE_RES)
    s.negatives = sum(1 for w in lower if w in NEGATIVE_TOKENS or w.endswith("n't"))
    s.double_neg = find_double_negatives(text)
    s.llm = find_phrases(text.lstrip(SENTENCE_START_CHARS + " "), LLM_RES)
    s.dashes = len(DASH_RE.findall(text))
    s.symbols = len(SYMBOL_RE.findall(text))
    return s


def extract_segments(raw: str) -> list[Segment]:
    raw = HTML_COMMENT_RE.sub(lambda m: "\n" * m.group(0).count("\n"), raw.translate(QUOTE_TRANSLATION))
    lines = raw.splitlines()
    start = 0
    if lines and lines[0].strip() == "---":
        for j in range(1, len(lines)):
            if lines[j].strip() in ("---", "..."):
                start = j + 1
                break

    segments: list[Segment] = []
    para: list[tuple[int, str]] = []
    fence = None

    def flush() -> None:
        if para:
            segments.append(Segment(para[0][0], "\n".join(t for _, t in para), "prose"))
            para.clear()

    for n in range(start, len(lines)):
        lineno = n + 1
        line = lines[n]
        if fence:
            if line.strip().startswith(fence):
                fence = None
            continue
        m = FENCE_RE.match(line)
        if m:
            flush()
            fence = m.group(1)
            continue
        line = QUOTE_RE.sub("", line)
        if not line.strip() or HR_RE.match(line):
            flush()
            continue
        m = HEADING_RE.match(line)
        if m:
            flush()
            segments.append(Segment(lineno, m.group(1), "heading"))
            continue
        if line.lstrip().startswith("|"):
            flush()
            if not TABLE_SEP_RE.match(line):
                for cell in line.strip().strip("|").split("|"):
                    if cell.strip():
                        segments.append(Segment(lineno, cell.strip(), "cell"))
            continue
        m = LIST_RE.match(line)
        if m:
            flush()
            line = line[m.end():]
        para.append((lineno, line.strip()))
    flush()
    return segments


def clean_inline(text: str) -> str:
    for regex, repl in INLINE_CLEANERS:
        text = regex.sub(repl, text)
    return text


def split_sentences(text: str) -> list[tuple[int, str]]:
    sentences = []
    start = 0
    for m in BOUNDARY_RE.finditer(text):
        nxt = text[m.end():m.end() + 1]
        if nxt and not (nxt.isupper() or nxt.isdigit() or nxt in SENTENCE_START_CHARS):
            continue
        prev = re.search(r"(\S+)$", text[start:m.start() + 1])
        if prev:
            token = prev.group(1).rstrip(".").lstrip("\"'([").lower()
            if token in ABBREVIATIONS or (len(token) == 1 and token.isalpha()):
                continue
        sentences.append((start, text[start:m.end()]))
        start = m.end()
    if text[start:].strip():
        sentences.append((start, text[start:]))
    return sentences


def analyze_document(path: Path) -> Report:
    raw = path.read_text(encoding="utf-8-sig", errors="replace")
    sentences: list[Sentence] = []
    blocks: list[list[Sentence]] = []
    for seg in extract_segments(raw):
        text = clean_inline(seg.text)
        if seg.kind != "prose":
            flat = " ".join(text.split())
            if WORD_RE.search(flat):
                sentences.append(analyze_sentence(seg.line, seg.kind, flat))
            continue
        block = []
        for offset, sent in split_sentences(text):
            flat = " ".join(sent.split())
            if WORD_RE.search(flat):
                block.append(analyze_sentence(seg.line + text.count("\n", 0, offset), "prose", flat))
        sentences.extend(block)
        if block:
            blocks.append(block)
    return Report(path, sentences, blocks)


def word_count(rep: Report) -> int:
    return sum(len(s.words) for s in rep.sentences)


def monotone_blocks(rep: Report) -> list[list[Sentence]]:
    result = []
    for block in rep.blocks:
        lengths = [len(s.words) for s in block]
        if len(lengths) >= MONOTONE_MIN_SENTENCES and max(lengths) - min(lengths) <= MONOTONE_SPREAD:
            result.append(block)
    return result


def dash_heavy_blocks(rep: Report) -> list[list[Sentence]]:
    return [b for b in rep.blocks if sum(s.dashes for s in b) >= 2]


def summarize(rep: Report) -> list[tuple[str, float, str]]:
    words = word_count(rep)
    lengths = [len(s.words) for s in rep.sentences if s.kind == "prose"]

    def total(fn) -> int:
        return sum(fn(s) for s in rep.sentences)

    def per100(n: int) -> float:
        return 100.0 * n / words if words else 0.0

    return [
        ("Words", words, "{:,.0f}"),
        ("Sentences", len(lengths), "{:,.0f}"),
        ("Mean sentence length", statistics.mean(lengths) if lengths else 0, "{:.1f}"),
        ("Sentence length std dev", statistics.pstdev(lengths) if len(lengths) > 1 else 0, "{:.1f}"),
        ("Longest sentence (words)", max(lengths, default=0), "{:.0f}"),
        (f"Sentences over {LONG_SENTENCE} words", sum(n > LONG_SENTENCE for n in lengths), "{:.0f}"),
        ("Prepositions per 100 words", per100(total(lambda s: len(s.preps))), "{:.1f}"),
        ('"To be" forms per 100 words', per100(total(lambda s: len(s.be))), "{:.1f}"),
        ("Nominalizations per 100 words", per100(total(lambda s: len(s.noms))), "{:.1f}"),
        ('"the <noun> of" patterns', total(lambda s: len(s.noun_of)), "{:.0f}"),
        ("Passive candidates", total(lambda s: s.passive), "{:.0f}"),
        ("Prepositional strings (3+)", total(lambda s: len(s.chains)), "{:.0f}"),
        ("There is / It is openers", total(lambda s: int(s.expletive)), "{:.0f}"),
        ("Lard phrases", total(lambda s: len(s.lard)), "{:.0f}"),
        ("Negatives", total(lambda s: s.negatives), "{:.0f}"),
        ("Double-negative candidates", total(lambda s: len(s.double_neg)), "{:.0f}"),
        ("Monotone paragraphs", len(monotone_blocks(rep)), "{:.0f}"),
        ("LLM tells", total(lambda s: len(s.llm)), "{:.0f}"),
        ("Em dashes", total(lambda s: s.dashes), "{:.0f}"),
        ("Paragraphs with 2+ em dashes", len(dash_heavy_blocks(rep)), "{:.0f}"),
        ("Arrows and emoji", total(lambda s: s.symbols), "{:.0f}"),
        ("Normative keywords (shall/must)", total(lambda s: len(NORMATIVE_RE.findall(s.text))), "{:.0f}"),
    ]


def markup(s: Sentence) -> str:
    def repl(m: re.Match) -> str:
        w = m.group(0)
        lw = w.lower()
        if lw in PREPOSITIONS:
            return f"[{w}]"
        if lw in BE_FORMS:
            return f"{{{w}}}"
        if is_nominalization(lw):
            return f"<{w}>"
        return w

    return WORD_RE.sub(repl, s.text)


def wrap(text: str, indent: str = "    ") -> str:
    return textwrap.fill(text, width=100, initial_indent=indent, subsequent_indent=indent)


# PowerShell often decodes piped UTF-8 with the OEM code page, so keep common punctuation ASCII.
ASCII_OUT = str.maketrans({"\u2014": "--", "\u2013": "-", "\u2026": "...", "\u2192": "->", "\u2190": "<-", "\u00a0": " "})


def emit(text: str = "") -> None:
    print(text.translate(ASCII_OUT))


def section(title: str, rows: list[str]) -> None:
    emit(f"\n== {title} ==")
    for row in rows or ["  (none)"]:
        emit(row)


def cmd_scan(args: argparse.Namespace) -> None:
    rep = analyze_document(Path(args.file))
    emit(f"== Prose audit: {rep.path} ==")
    emit("(Excludes front matter, code, inline code, URLs, and HTML.)")
    for label, value, fmt in summarize(rep):
        emit(f"  {label:<32} {fmt.format(value)}")

    emit("\nMarkup legend: [preposition]  {to-be form}  <nominalization>  ('to' includes infinitives)")
    worst = sorted((s for s in rep.sentences if s.score > 0), key=lambda s: s.score, reverse=True)[: args.top]
    rows = []
    for s in worst:
        rows.append(f"\nL{s.line} ({len(s.words)} words, score {s.score:.0f})")
        rows.append(wrap(markup(s)))
        for chain in s.chains:
            rows.append(f"    - prepositional string: {' > '.join(chain)}")
        if s.passive:
            rows.append(f"    - passive candidates: {s.passive}")
        for noun in s.noun_of:
            rows.append(f"    - buried verb: 'the {noun} of'")
        for phrase, suggestion in s.lard:
            rows.append(f"    - lard: '{phrase}' -> {suggestion}")
        if s.expletive:
            rows.append("    - slow start: There is / It is opener")
        for hit in s.double_neg:
            rows.append(f"    - double negative: '{hit}'")
        for phrase, suggestion in s.llm:
            rows.append(f"    - LLM tell: '{phrase}' -> {suggestion}")
        if s.dashes:
            rows.append(f"    - em dashes: {s.dashes}")
        if s.symbols:
            rows.append(f"    - arrows/emoji: {s.symbols}")
    section(f"Top {len(worst)} sentences to revise", rows)

    section(
        "Lard phrases",
        [f"  L{s.line}: '{p}' -> {sug}" for s in rep.sentences for p, sug in s.lard],
    )
    section(
        "Double-negative candidates",
        [f"  L{s.line}: '{hit}'" for s in rep.sentences for hit in s.double_neg],
    )
    section(
        "There is / It is openers",
        [f"  L{s.line}: {' '.join(s.text.split()[:10])} ..." for s in rep.sentences if s.expletive],
    )
    section(
        f"Long sentences (> {LONG_SENTENCE} words)",
        [f"  L{s.line}: {len(s.words)} words" for s in rep.sentences if s.kind == "prose" and len(s.words) > LONG_SENTENCE],
    )
    section(
        "Rhythm: monotone paragraphs (sentence lengths)",
        [f"  L{b[0].line}: {' / '.join(str(len(s.words)) for s in b)}" for b in monotone_blocks(rep)],
    )
    section(
        "LLM tells",
        [f"  L{s.line}: '{p}' -> {sug}" for s in rep.sentences for p, sug in s.llm],
    )
    section(
        "Paragraphs with 2+ em dashes (aim for one at most)",
        [f"  L{b[0].line}: {sum(s.dashes for s in b)}" for b in dash_heavy_blocks(rep)],
    )
    section(
        "Arrows and emoji",
        [f"  L{s.line}: {s.symbols}" for s in rep.sentences if s.symbols],
    )


def cmd_compare(args: argparse.Namespace) -> None:
    before = analyze_document(Path(args.original))
    after = analyze_document(Path(args.revised))
    wb, wa = word_count(before), word_count(after)
    lf = 100.0 * (wb - wa) / wb if wb else 0.0
    delta = f"{wb - wa:,} cut" if wb >= wa else f"{wa - wb:,} added"
    print(f"**Lard Factor: {lf:.1f}%** ({wb:,} -> {wa:,} words, {delta})\n")
    print("| Metric | Before | After |")
    print("|--------|-------:|------:|")
    for (label, vb, fmt), (_, va, _) in zip(summarize(before), summarize(after)):
        print(f"| {label} | {fmt.format(vb)} | {fmt.format(va)} |")


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    scan = sub.add_parser("scan", help="Paramedic Method markup and wordiness metrics")
    scan.add_argument("file")
    scan.add_argument("--top", type=int, default=15, help="number of worst sentences to show (default 15)")
    scan.set_defaults(func=cmd_scan)
    compare = sub.add_parser("compare", help="Lard Factor and before/after metrics")
    compare.add_argument("original")
    compare.add_argument("revised")
    compare.set_defaults(func=cmd_compare)
    args = parser.parse_args()
    for name in ("file", "original", "revised"):
        path = getattr(args, name, None)
        if path and not Path(path).is_file():
            parser.error(f"not a file: {path}")
    args.func(args)
    return 0


if __name__ == "__main__":
    sys.exit(main())

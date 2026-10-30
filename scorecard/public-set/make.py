"""Make the public document set for the page-layout scorecard.

    py -3.13 -I scorecard/public-set/make.py        (run from the repo root)

Writes eleven invented Word documents into scorecard/public-set/docs/, plus set.json next to this file.
Every word of text is invented here from a built-in vocabulary; nothing is read from anywhere else.
The output is deterministic: the same seed gives byte-identical files, and no timestamps are written.
Only the standard library is used.
"""
import json
import os
import random
import re
import struct
import sys
import zipfile
import zlib
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(HERE, 'docs')
SEED = 20261029

# --------------------------------------------------------------------------------------------------
# Vocabulary (invented prose about an invented harbour rescue service)
# --------------------------------------------------------------------------------------------------

NOUNS = """
crew boat vessel lifeboat launch tender dinghy skiff harbour pier jetty quay wharf slipway mooring berth
buoy beacon channel fairway bar reef shoal sandbank tide current swell wave surf breaker wind gust squall
gale storm fog mist haze rain drizzle spray visibility horizon shoreline headland cove inlet estuary lagoon
marsh mudflat cliff dune beach shingle station boathouse watchtower lookout radio transmitter receiver
aerial antenna signal frequency callsign message report log logbook entry record register roster rota
shift watch duty officer coxswain helmsman navigator mechanic engineer medic swimmer volunteer trainee
instructor supervisor coordinator controller inspector auditor casualty survivor passenger skipper owner
angler diver kayaker child adult family group team unit squad section division department service agency
authority council board committee meeting briefing debrief review inspection audit exercise drill
rehearsal training course module lesson assessment examination certificate qualification licence permit
approval authorisation clearance notice warning alert alarm siren flare rocket lamp torch lantern
lifejacket lifebuoy lifeline harness helmet drysuit wetsuit glove boot stretcher splint bandage dressing
blanket oxygen kit pack bag box locker cabinet shelf rack hook line rope cable chain anchor winch davit
crane hoist pulley block shackle cleat bollard fender hull deck bow stern keel rudder tiller wheel
throttle engine motor propeller gearbox fuel diesel petrol oil coolant filter pump bilge valve hose nozzle
tank battery charger fuse switch circuit wire connector gauge meter dial indicator display screen console
panel keypad button lever handle latch hatch door window porthole ladder step rail railing gangway ramp
bridge cabin wheelhouse galley bunk workshop store yard compound gate fence road lane track path route
passage crossing journey voyage trip callout recovery search rescue tow escort transfer evacuation landing
pickup handover response incident emergency accident collision grounding capsize flooding fire leak
breakdown failure fault defect damage repair maintenance overhaul replacement upgrade modification
procedure process method practice standard rule regulation requirement condition limit threshold margin
tolerance allowance exception exemption waiver policy guideline instruction direction order command
request question answer decision outcome result finding conclusion recommendation action task job
responsibility role position post rank grade level stage phase sequence list table chart diagram map plan
drawing sketch photograph picture image copy original version edition revision draft document page
chapter clause paragraph item point note remark comment observation detail summary overview introduction
background purpose scope definition term word name number code reference identifier label tag mark symbol
sign colour shape size length width height depth weight volume capacity speed rate hour minute second day
night morning afternoon evening week month season year date period interval delay pause break start
finish beginning end middle centre side edge corner top bottom front back direction bearing heading
course distance range area zone region district sector quadrant grid square circle triangle water sea
ocean river lake pond pool stream creek drain outfall pipe culvert weather climate temperature pressure
humidity forecast outlook hazard risk danger threat safety security protection cover shelter refuge help
aid assistance support backup relief cost price budget fund grant donation payment invoice receipt account
balance total sum amount quantity count figure estimate average maximum minimum percentage proportion
share part piece portion fraction half quarter third follow-up call-out check-list run-down roll-call
half-tide high-water low-water set-up hand-over changeover trailer tractor van lorry vehicle keel-band
paddle oar sail mast boom rigging tarpaulin cradle chock strap buckle clip pin bolt nut washer bracket
hinge seal gasket bearing shaft coupling belt sensor transducer compass chartplotter sounder radar plotter
binoculars whistle horn bell flag pennant marker stake post pole tower mast ledge shelf platform pontoon
raft float cushion seat bench cupboard drawer tray tub bucket bottle can carton sack crate pallet
coupon ticket voucher form sheet card folder binder envelope letter memo circular bulletin newsletter
poster leaflet manual handbook guide primer glossary appendix index footnote caption heading subheading
margin column row cell border frame layout outline template pattern model sample specimen example
instance case matter issue problem difficulty obstacle setback mistake error slip lapse oversight
remedy solution fix workaround improvement gain loss saving waste surplus shortage deficit reserve stock
supply inventory asset tool spanner wrench screwdriver hammer mallet chisel saw drill file rasp brush
sponge cloth rag towel soap detergent grease wax paint varnish primer sealant adhesive tape cord twine
mesh net netting gauze foil sheeting board plank beam joist rafter purlin batten tile slate brick stone
gravel sand silt clay mud peat turf moss weed kelp wrack driftwood debris litter rubbish wreckage
""".split()

ADJECTIVES = """
small large long short high low wide narrow deep shallow heavy light strong weak fast slow quick steady
calm rough smooth flat steep sharp blunt hot cold warm cool wet dry damp clean dirty clear cloudy bright
dark pale loud quiet early late new old young recent previous current final first last next other same
different similar separate single double full empty open closed ready busy free available suitable proper
correct wrong safe dangerous careful careless simple complex easy hard difficult useful useless important
minor major main primary secondary general special specific particular usual normal ordinary regular
routine standard basic formal informal official local regional national coastal inshore offshore tidal
seasonal annual monthly weekly daily hourly temporary permanent immediate urgent serious severe moderate
slight gradual sudden rapid brief lengthy complete partial whole entire total exact approximate accurate
reliable faulty broken damaged missing spare extra additional further remaining outstanding overdue likely
unlikely possible probable certain obvious visible hidden audible silent written verbal printed electronic
manual automatic mechanical electrical hydraulic portable fixed mobile inflatable rigid wooden metal
plastic rubber canvas orange yellow red white blue green grey black northern southern eastern western
upper lower inner outer nearby distant adjacent opposite parallel vertical horizontal circular square
rectangular twenty-four-hour well-known long-standing up-to-date hands-on day-to-day self-contained
water-resistant purpose-built low-lying hard-wearing part-time full-time half-hourly all-weather hand-held
second-hand long-term short-term one-off two-way three-way well-worn sea-going shore-based
thorough careful prompt alert aware fit tired cheerful patient firm gentle loose tight slack taut
level uneven sloping curved straight bent twisted cracked rusty worn fresh stale sweet bitter salty
brackish muddy sandy rocky stony grassy marshy misty foggy rainy windy stormy sunny overcast humid
frosty icy snowy mild harsh bleak remote isolated crowded sparse dense thick thin fine coarse soft
brittle flexible sturdy robust fragile delicate sound seaworthy roadworthy valid expired lapsed
confidential public private internal external verbal pending active inactive dormant vacant occupied
""".split()

VERBS = """
check inspect examine test record note report log mark sign date file store keep hold carry lift lower
raise move shift turn pull push drag haul tow launch recover return repair replace remove fit attach
fasten secure release loosen tighten open close lock unlock start stop run operate use handle control
steer drive sail row paddle swim wade climb descend cross enter leave approach reach pass follow lead
guide direct instruct train teach learn practise rehearse prepare plan arrange organise manage supervise
monitor watch observe see look listen hear call answer reply speak say tell ask confirm agree refuse
accept reject approve allow permit require need want expect intend decide choose select consider assess
judge measure count estimate calculate compare match apply include exclude cover protect warn alert
notify inform advise remind explain describe show display present demonstrate prove find locate identify
recognise discover search seek lose miss forget remember know understand believe think suppose assume
hope wish try attempt fail succeed complete finish begin continue resume suspend cancel postpone delay
wait stay remain rest sleep wake eat drink wash dry clean tidy sort pack unpack load unload fill empty
drain pump pour spray wipe cut tie knot splice coil wind unwind hang mount build make produce provide
supply issue deliver send receive collect gather bring take give pay buy sell hire borrow lend share
divide split join connect link separate mix combine change adjust improve reduce increase extend limit
restrict prevent avoid cause affect damage break bend crack wear rust leak sink float drift roll pitch
heel ground strand capsize flood burn smoke signal radio telephone contact summon dispatch deploy
assemble muster disband relieve rotate rota assign allocate reserve book cancel confirm renew revoke
suspend extend amend revise update review audit certify verify validate inspect sample weigh gauge
calibrate service overhaul refit refuel recharge restock replenish rinse scrub polish grease lubricate
tension slacken reef furl hoist stow lash rig unrig moor unmoor berth beach refloat right bail drain
flush vent seal patch weld rivet bolt screw glue nail stitch sew mend darn trim shorten lengthen widen
narrow deepen dredge survey chart plot track trace sight range fix log transmit relay acknowledge
read write print copy scan post mail hand pass forward deliver return attend visit tour patrol scout
""".split()

ADVERBS = """
quickly slowly carefully properly firmly gently immediately promptly regularly frequently rarely seldom
occasionally usually normally generally typically mostly partly fully entirely completely nearly almost
roughly exactly clearly plainly simply easily safely securely directly briefly shortly recently previously
formerly currently presently eventually finally lastly firstly secondly thirdly separately together apart
aside ahead behind inside outside overboard ashore aboard afloat underway overnight meanwhile otherwise
therefore however moreover nevertheless instead anyway perhaps probably certainly possibly obviously
apparently especially particularly notably mainly chiefly largely hardly barely scarcely twice once
again already still yet soon later earlier sometimes often always never elsewhere nearby upstream
downstream inshore offshore seaward landward northward southward eastward westward
""".split()

FUNCTION_WORDS = """
the a an this that these those each every any some all no both either neither such its their our his
her of in on at by for from to with without within under over into onto through across along around
between among before after during until since above below beyond near against toward towards about
and or but nor so yet if when while because although though unless whether as than which who where
how why not it they we you one is are was were be been being has have had do does did will would shall
should may might must can could there here then also only still just even quite rather very too
""".split()

PLACES = """
Greyhaven|Saltmarsh Point|Tallow Reach|Cobble Strand|Larkspur Bay|Gull Rock|Ninefold Bank|Crane Water
Varrow Head|Scarrow Flats|Brinemouth|Kestrel Bank|Pennick Sands|Ashgate Pier|Mallow Creek|Old Haven
Ferrow Light|Wrack Island|Dunmere|Shale Point|Corrin Sound|Thistle Rock|Marrow Spit|Low Holm|Ebb Gate
Fennel Cove|Sorrel Bank|Tern Island|Haddock Reach|Lantern Quay|Tidewell|Blackstone Ledge|Rushy Inlet
""".replace('\n', '|').split('|')
PLACES = [p for p in PLACES if p]

ROLES = """
the duty officer|the coxswain|the helmsman|the mechanic|the station manager|the watch leader
the training officer|the launch authority|the shore crew|the boat crew|the duty medic|the lookout
the relief crew|the inspector|the harbour master|the second coxswain|the navigator|the crew
""".replace('\n', '|').split('|')
ROLES = [p for p in ROLES if p]

UNITS = 'metres minutes knots litres hours degrees kilograms seconds miles tonnes'.split()
MODALS = 'must should may can will shall might could'.split()
PREPS = ('in on at by for from with without within under over into through across along around between '
         'before after during until above below beyond near against towards about').split()
SAYS = 'said reported noted confirmed recorded observed explained added wrote agreed'.split()
REFS = ('section 2.1|section 4.2|section 7.3|chapter 3|chapter 9|table 2|table 5|the appendix|rule 12|'
        'rule 31|form B|figure 4|the glossary|part two|clause 8.4|note 3').split('|')
DAYS = 'Monday Tuesday Wednesday Thursday Friday Saturday Sunday'.split()
NUMBER_WORDS = 'two three four five six seven eight nine ten twelve fifteen twenty thirty forty fifty'.split()

IRREGULAR_PAST = {
    'begin': ('began', 'begun'), 'bring': ('brought', 'brought'), 'build': ('built', 'built'),
    'buy': ('bought', 'bought'), 'choose': ('chose', 'chosen'), 'cut': ('cut', 'cut'),
    'drink': ('drank', 'drunk'), 'drive': ('drove', 'driven'), 'eat': ('ate', 'eaten'),
    'find': ('found', 'found'), 'fit': ('fitted', 'fitted'), 'forget': ('forgot', 'forgotten'),
    'give': ('gave', 'given'), 'hang': ('hung', 'hung'), 'hear': ('heard', 'heard'),
    'hold': ('held', 'held'), 'keep': ('kept', 'kept'), 'know': ('knew', 'known'),
    'lead': ('led', 'led'), 'leave': ('left', 'left'), 'lend': ('lent', 'lent'),
    'lose': ('lost', 'lost'), 'make': ('made', 'made'), 'pay': ('paid', 'paid'),
    'read': ('read', 'read'), 'run': ('ran', 'run'), 'say': ('said', 'said'), 'see': ('saw', 'seen'),
    'seek': ('sought', 'sought'), 'sell': ('sold', 'sold'), 'send': ('sent', 'sent'),
    'sink': ('sank', 'sunk'), 'sleep': ('slept', 'slept'), 'speak': ('spoke', 'spoken'),
    'split': ('split', 'split'), 'swim': ('swam', 'swum'), 'take': ('took', 'taken'),
    'teach': ('taught', 'taught'), 'tell': ('told', 'told'), 'think': ('thought', 'thought'),
    'understand': ('understood', 'understood'), 'wake': ('woke', 'woken'), 'wear': ('wore', 'worn'),
    'wind': ('wound', 'wound'), 'write': ('wrote', 'written'), 'break': ('broke', 'broken'),
    'bend': ('bent', 'bent'), 'bind': ('bound', 'bound'), 'draw': ('drew', 'drawn'),
    'spread': ('spread', 'spread'), 'deal': ('dealt', 'dealt'), 'dig': ('dug', 'dug'),
    'hit': ('hit', 'hit'), 'let': ('let', 'let'), 'lay': ('laid', 'laid'), 'light': ('lit', 'lit'),
    'feed': ('fed', 'fed'), 'hide': ('hid', 'hidden'), 'blow': ('blew', 'blown'),
    'grow': ('grew', 'grown'), 'fly': ('flew', 'flown'), 'freeze': ('froze', 'frozen'),
    'stand': ('stood', 'stood'), 'sit': ('sat', 'sat'), 'strike': ('struck', 'struck'),
    'stick': ('stuck', 'stuck'), 'shut': ('shut', 'shut'), 'set': ('set', 'set'), 'put': ('put', 'put'),
    'ride': ('rode', 'ridden'), 'rise': ('rose', 'risen'), 'ring': ('rang', 'rung'),
    'meet': ('met', 'met'), 'mean': ('meant', 'meant'), 'go': ('went', 'gone'), 'get': ('got', 'got'),
    'come': ('came', 'come'), 'do': ('did', 'done'), 'have': ('had', 'had'), 'catch': ('caught', 'caught'),
    'fall': ('fell', 'fallen'), 'feel': ('felt', 'felt'), 'spend': ('spent', 'spent'), 'win': ('won', 'won'),
    'throw': ('threw', 'thrown'), 'tie': ('tied', 'tied'), 'shake': ('shook', 'shaken'),
    'cost': ('cost', 'cost'), 'shine': ('shone', 'shone'), 'right': ('righted', 'righted'),
    'rota': ('rotaed', 'rotaed'), 'radio': ('radioed', 'radioed'), 'ski': ('skied', 'skied'),
}
IRREGULAR_PLURAL = {
    'child': 'children', 'shelf': 'shelves', 'half': 'halves', 'wharf': 'wharves', 'glossary': 'glossaries',
    'index': 'indexes', 'appendix': 'appendices', 'crew': 'crews', 'series': 'series', 'oxygen': 'oxygen',
    'fuel': 'fuel', 'diesel': 'diesel', 'petrol': 'petrol', 'oil': 'oil', 'water': 'water', 'weather': 'weather',
    'sand': 'sand', 'mud': 'mud', 'kelp': 'kelp', 'wrack': 'wrack', 'debris': 'debris', 'litter': 'litter',
    'rubbish': 'rubbish', 'wreckage': 'wreckage', 'gravel': 'gravel', 'silt': 'silt', 'clay': 'clay',
    'peat': 'peat', 'turf': 'turf', 'moss': 'moss', 'grease': 'grease', 'wax': 'waxes', 'paint': 'paint',
    'rigging': 'rigging', 'netting': 'netting', 'gauze': 'gauze', 'foil': 'foil', 'sheeting': 'sheeting',
    'binoculars': 'binoculars', 'driftwood': 'driftwood', 'coolant': 'coolant', 'safety': 'safety',
    'visibility': 'visibility', 'maintenance': 'maintenance', 'training': 'training', 'help': 'help',
    'assistance': 'assistance', 'support': 'support', 'shingle': 'shingle', 'haze': 'haze', 'mist': 'mist',
    'fog': 'fog', 'rain': 'rain', 'drizzle': 'drizzle', 'spray': 'spray', 'surf': 'surf', 'swell': 'swell',
    'soap': 'soap', 'detergent': 'detergent', 'varnish': 'varnish', 'sealant': 'sealant', 'adhesive': 'adhesive',
    'tape': 'tape', 'cord': 'cord', 'twine': 'twine', 'mesh': 'mesh', 'stock': 'stock', 'supply': 'supplies',
}

VOWELS = 'aeiou'


def plural(noun):
    if noun in IRREGULAR_PLURAL:
        return IRREGULAR_PLURAL[noun]
    if noun.endswith('y') and noun[-2] not in VOWELS:
        return noun[:-1] + 'ies'
    if noun.endswith(('s', 'x', 'z', 'ch', 'sh')):
        return noun + 'es'
    return noun + 's'


def past(verb):
    if verb in IRREGULAR_PAST:
        return IRREGULAR_PAST[verb][0]
    if verb.endswith('e'):
        return verb + 'd'
    if verb.endswith('y') and verb[-2] not in VOWELS:
        return verb[:-1] + 'ied'
    if (len(verb) <= 4 and verb[-1] not in 'wxy' and verb[-1] not in VOWELS
            and verb[-2] in VOWELS and verb[-3] not in VOWELS):
        return verb + verb[-1] + 'ed'
    return verb + 'ed'


def participle(verb):
    if verb in IRREGULAR_PAST:
        return IRREGULAR_PAST[verb][1]
    return past(verb)


def third(verb):
    if verb.endswith('y') and verb[-2] not in VOWELS:
        return verb[:-1] + 'ies'
    if verb.endswith(('s', 'x', 'z', 'ch', 'sh', 'o')):
        return verb + 'es'
    return verb + 's'


def article(word):
    return ('an ' if word[0] in VOWELS else 'a ') + word


def cap(text):
    return text[0].upper() + text[1:]


class Prose:
    """Invented English sentences, paragraphs and headings from the vocabulary above."""

    def __init__(self, rng):
        self.r = rng

    def noun_phrase(self, want_plural=None):
        r = self.r
        k = r.random()
        if k < 0.12:
            return r.choice(ROLES), False
        if k < 0.18:
            return r.choice(PLACES), False
        noun = r.choice(NOUNS)
        is_plural = (r.random() < 0.3) if want_plural is None else want_plural
        k = r.random()
        n_adj = 0 if k < 0.55 else (1 if k < 0.9 else 2)
        adjs = r.sample(ADJECTIVES, n_adj)
        core = ' '.join(adjs + [plural(noun) if is_plural else noun])
        if is_plural:
            det = r.choice(['the', 'the', 'the', '', 'these', 'all the', 'most', 'several', 'two', 'three',
                            'the remaining', 'any', 'both', 'some', 'the other'])
            phrase = (det + ' ' + core) if det else core
        else:
            det = r.choice(['the', 'the', 'the', 'the', 'a', 'a', 'this', 'that', 'each', 'every', 'one',
                            'its', 'their', 'our', 'another', 'no', 'the same', 'the first', 'the next'])
            if det == 'a':
                phrase = article(core)
            else:
                phrase = det + ' ' + core
        if r.random() < 0.15:
            tail, _ = self.noun_phrase(False)
            phrase += ' of ' + tail
        return phrase, is_plural

    def pp(self):
        prep = self.r.choice(PREPS)
        np, _ = self.noun_phrase()
        return prep + ' ' + np

    def number(self):
        r = self.r
        k = r.random()
        if k < 0.35:
            return str(r.randint(2, 99))
        if k < 0.55:
            return r.choice(NUMBER_WORDS)
        if k < 0.7:
            return str(r.randint(100, 950))
        if k < 0.8:
            return '{:,}'.format(r.randint(1000, 24000))
        if k < 0.9:
            return '%d.%d' % (r.randint(0, 12), r.randint(1, 9))
        return str(r.randint(1961, 2025))

    def verb_phrase(self, tense, is_plural=False):
        r = self.r
        verb = r.choice(VERBS)
        if tense == 'past':
            form = past(verb)
        elif tense == 'present':
            form = verb if is_plural else third(verb)
        elif tense == 'modal':
            form = r.choice(MODALS) + ' ' + verb
        else:
            form = verb
        if r.random() < 0.2:
            form = r.choice(ADVERBS) + ' ' + form
        k = r.random()
        if k < 0.6:
            obj, _ = self.noun_phrase()
            form += ' ' + obj
        elif k < 0.8:
            form += ' ' + self.pp()
        else:
            form += ' ' + r.choice(ADVERBS)
        return form

    def clause(self, tense='past'):
        subj, pl = self.noun_phrase()
        return subj + ' ' + self.verb_phrase(tense, pl)

    def sentence_text(self):
        r = self.r
        t = r.randrange(19)
        if t == 0:
            s = '%s %s.' % (self.clause('past'), self.pp())
        elif t == 1:
            s = '%s, %s, and %s.' % (cap(r.choice(ADVERBS)), self.clause('past'), self.clause('past'))
        elif t == 2:
            np, _ = self.noun_phrase()
            s = 'If %s, %s must %s %s %s.' % (self.clause('present'), r.choice(ROLES), r.choice(VERBS), np, self.pp())
        elif t == 3:
            np, _ = self.noun_phrase()
            s = '%s should be %s %s before %s.' % (cap(np), participle(r.choice(VERBS)), self.pp(), self.clause('present'))
        elif t == 4:
            s = '%s %s that %s; %s.' % (cap(r.choice(ROLES)), r.choice(SAYS), self.clause('past'), self.clause('past'))
        elif t == 5:
            np, _ = self.noun_phrase()
            s = '%s (%s %s) %s %s %s.' % (cap(np), self.number(), r.choice(UNITS), past(r.choice(VERBS)), self.pp(), self.pp())
        elif t == 6:
            np, pl = self.noun_phrase()
            s = 'When %s %s %s, %s %s.' % (np, 'are' if pl else 'is', r.choice(ADJECTIVES), self.clause('present'), r.choice(ADVERBS))
        elif t == 7:
            s = 'In %d %s at %s %s %s %s %s.' % (r.randint(1961, 2025), r.choice(ROLES), r.choice(PLACES),
                                                 past(r.choice(VERBS)), self.number(), plural(r.choice(NOUNS)), self.pp())
        elif t == 8:
            np, pl = self.noun_phrase()
            s = '%s %s %s and %s, but %s.' % (cap(np), 'are' if pl else 'is', r.choice(ADJECTIVES), r.choice(ADJECTIVES), self.clause('present'))
        elif t == 9:
            np, _ = self.noun_phrase()
            s = 'No %s may be %s %s unless %s has been %s %s.' % (r.choice(NOUNS), participle(r.choice(VERBS)), self.pp(),
                                                                   np, participle(r.choice(VERBS)), self.pp())
        elif t == 10:
            np, _ = self.noun_phrase()
            s = 'Each %s %s %s %s %s.' % (r.choice(NOUNS), third(r.choice(VERBS)), np, r.choice(ADVERBS), self.pp())
        elif t == 11:
            s = 'Between %s and %s, %s %s.' % (r.choice(PLACES), r.choice(PLACES), self.clause('past'), self.pp())
        elif t == 12:
            s = '%s %s; otherwise %s.' % (cap(self.clause('modal')), self.pp(), self.clause('modal'))
        elif t == 13:
            s = '%s %s, although %s.' % (cap(r.choice(ADVERBS)), self.clause('past'), self.clause('past'))
        elif t == 14:
            np, _ = self.noun_phrase()
            s = 'The %s of %s %s %s (see %s).' % (r.choice(NOUNS), np, past(r.choice(VERBS)), self.pp(), r.choice(REFS))
        elif t == 15:
            s = '%s, %s, and %s.' % (cap(self.clause('past')), self.clause('past'), self.clause('past'))
        elif t == 16:
            np, _ = self.noun_phrase()
            s = 'At %02d%02d on %s, %s %s after %s %s %s.' % (r.randint(0, 23), r.choice([0, 15, 30, 45]), r.choice(DAYS),
                                                               self.clause('past'), self.pp(), np, past(r.choice(VERBS)), self.pp())
        elif t == 17:
            np, _ = self.noun_phrase()
            obj, _ = self.noun_phrase()
            s = '%s %s %s %s %s %s, which %s %s.' % (cap(np), past(r.choice(VERBS)), self.number(), r.choice(UNITS),
                                                     r.choice(ADVERBS), self.pp(), past(r.choice(VERBS)), obj)
        else:
            s = '%s.' % cap(self.clause('past'))
        return s

    def sentence(self, lo=6, hi=35):
        while True:
            s = self.sentence_text()
            n = len(s.split(' '))
            if lo <= n <= hi:
                return s

    def paragraph(self, lo=2, hi=9):
        n = self.r.randint(lo, hi)
        return ' '.join(self.sentence() for _ in range(n))

    def short(self, lo=1, hi=2):
        n = self.r.randint(lo, hi)
        return ' '.join(self.sentence(6, 18) for _ in range(n))

    def title(self, lo=2, hi=5):
        r = self.r
        n = r.randint(lo, hi)
        words = []
        for i in range(n):
            k = r.random()
            if k < 0.45:
                words.append(r.choice(NOUNS))
            elif k < 0.75:
                words.append(r.choice(ADJECTIVES))
            elif k < 0.9:
                words.append(plural(r.choice(NOUNS)))
            else:
                words.append(r.choice(['and', 'of', 'for', 'at', 'on', 'in']))
        out = []
        for i, w in enumerate(words):
            if i == 0 or w not in ('and', 'of', 'for', 'at', 'on', 'in'):
                w = cap(w)
            out.append(w)
        return ' '.join(out)

    def cell(self):
        r = self.r
        k = r.random()
        if k < 0.3:
            return '%s %s' % (self.number(), r.choice(UNITS))
        if k < 0.6:
            np, _ = self.noun_phrase()
            return cap(np)
        if k < 0.8:
            return cap(r.choice(ADJECTIVES))
        return self.sentence(6, 14)

    def label(self):
        r = self.r
        return cap(r.choice(ADJECTIVES) + ' ' + r.choice(NOUNS))


# --------------------------------------------------------------------------------------------------
# PNG pictures made from scratch
# --------------------------------------------------------------------------------------------------

def png_chunk(kind, data):
    body = kind + data
    return struct.pack('>I', len(data)) + body + struct.pack('>I', zlib.crc32(body) & 0xffffffff)


def make_png(width, height, pixel):
    """pixel(x, y) -> (r, g, b); an 8-bit RGB PNG with no interlace."""
    rows = []
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            row.extend(pixel(x, y))
        rows.append(bytes(row))
    raw = b''.join(rows)
    ihdr = struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0)
    return (b'\x89PNG\r\n\x1a\n' + png_chunk(b'IHDR', ihdr)
            + png_chunk(b'IDAT', zlib.compress(raw, 9)) + png_chunk(b'IEND', b''))


def picture_bytes(kind, w=200, h=120):
    if kind == 0:
        return make_png(w, h, lambda x, y: (200, 60, 40) if (x // 25 + y // 24) % 2 == 0 else (240, 200, 60))
    if kind == 1:
        return make_png(w, h, lambda x, y: (x * 255 // (w - 1), 80, 255 - x * 255 // (w - 1)))
    if kind == 2:
        return make_png(w, h, lambda x, y: (30, 90 + y * 160 // (h - 1), 150))
    if kind == 3:
        return make_png(w, h, lambda x, y: (250, 250, 250) if 20 <= x < w - 20 and 20 <= y < h - 20 else (20, 70, 130))
    return make_png(w, h, lambda x, y: ((x * 7 + y * 3) % 256, (x + y) % 256, 120))


# --------------------------------------------------------------------------------------------------
# WordprocessingML writer
# --------------------------------------------------------------------------------------------------

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
NS_WP14 = 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
NS_WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
NS_V = 'urn:schemas-microsoft-com:vml'
NS_O = 'urn:schemas-microsoft-com:office:office'
NS_W10 = 'urn:schemas-microsoft-com:office:word'

XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

ROOT_NS = (' xmlns:w="%s" xmlns:r="%s" xmlns:wp="%s" xmlns:a="%s" xmlns:pic="%s" xmlns:mc="%s" xmlns:w14="%s"'
           ' xmlns:wp14="%s" xmlns:wps="%s" xmlns:v="%s" xmlns:o="%s" xmlns:w10="%s" mc:Ignorable="w14 wp14"'
           % (NS_W, NS_R, NS_WP, NS_A, NS_PIC, NS_MC, NS_W14, NS_WP14, NS_WPS, NS_V, NS_O, NS_W10))

PAGE_W, PAGE_H, MARGIN = 12240, 15840, 1440
TEXT_W = PAGE_W - 2 * MARGIN   # 9360 twips, 6.5 inches

REL_TYPES = {
    'styles': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
    'numbering': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
    'settings': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
    'header': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
    'footer': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
    'image': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    'theme': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
}


def esc(text):
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def run(text, rpr=''):
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (('<w:rPr>%s</w:rPr>' % rpr) if rpr else '', esc(text))


def field(instr):
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> %s </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>1</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>' % instr)


class Doc:
    """One document: body chunks are collected, then the package is written."""

    def __init__(self, name, title, font, half_points, rng):
        self.name, self.title, self.font, self.sz = name, title, font, half_points
        self.r = rng
        self.prose = Prose(rng)
        self.body = []
        self.pars = []          # (kind, text) for every paragraph, document order; text boxes counted once
        self.used_ids = set()
        self.media = []         # (filename, bytes)
        self.nums = []          # (numId, abstractNumId)
        self.docpr = 0
        self.words = 0
        self.header_text = title
        self.header_align = 'right'
        self.theme = False          # write a theme part; the defaults then name theme fonts, as Word's own files do
        self.normal_ppr = ''        # paragraph properties on the Normal style
        self.extra_styles = ''      # more styles, appended to styles.xml
        self.round_bullets = False  # a third numbering definition: Symbol bullets with Word's hint attributes
        self.margins = (MARGIN, MARGIN, MARGIN, MARGIN, 720, 720)   # top, right, bottom, left, header, footer
        self.final_sect = None      # the section properties that close the document, when not the default

    # ---- ids ----
    def para_id(self):
        while True:
            v = self.r.getrandbits(31)
            if v and v not in self.used_ids:
                self.used_ids.add(v)
                return '%08X' % v

    def p_open(self):
        return '<w:p w14:paraId="%s" w14:textId="77777777">' % self.para_id()

    # ---- paragraphs ----
    def runs_for(self, text):
        """Mostly one run; now and then a word in italics or bold splits the paragraph into three runs."""
        words = text.split(' ')
        if len(words) > 8 and self.r.random() < 0.12:
            i = self.r.randint(2, len(words) - 3)
            rpr = '<w:i/><w:iCs/>' if self.r.random() < 0.6 else '<w:b/><w:bCs/>'
            return (run(' '.join(words[:i]) + ' ') + run(words[i], rpr) + run(' ' + ' '.join(words[i + 1:])))
        return run(text)

    def para_xml(self, text, style=None, num=None, lvl=0, pbdr='', spacing='', jc=None, rpr='', sect='',
                 keep_next=False, runs=None, pre='', ind=''):
        """pre: properties that come before spacing (kinsoku and the like); ind: a w:ind element."""
        ppr = ''
        if style:
            ppr += '<w:pStyle w:val="%s"/>' % style
        if keep_next:
            ppr += '<w:keepNext/>'
        if num is not None:
            ppr += '<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>' % (lvl, num)
        ppr += pbdr + pre + spacing + ind
        if jc:
            ppr += '<w:jc w:val="%s"/>' % jc
        if rpr:
            ppr += '<w:rPr>%s</w:rPr>' % rpr
        ppr += sect
        xml = self.p_open()
        if ppr:
            xml += '<w:pPr>%s</w:pPr>' % ppr
        if runs is not None:
            xml += runs
        elif text:
            xml += self.runs_for(text) if not rpr else run(text, rpr)
        return xml + '</w:p>'

    def p(self, text, kind='body', **kw):
        self.body.append(self.para_xml(text, **kw))
        self.pars.append((kind, text))
        self.words += len(text.split()) if text else 0

    def heading(self, level, text):
        self.p(text, kind='heading', style='Heading%d' % level)

    def page_break(self):
        self.body.append(self.p_open() + '<w:r><w:br w:type="page"/></w:r></w:p>')
        self.pars.append(('break', ''))

    def empty(self):
        self.p('', kind='empty')

    # ---- lists ----
    def new_num(self, abstract):
        if abstract == 2 and not self.round_bullets:
            raise SystemExit('%s: numbering 2 needs round_bullets' % self.name)
        num_id = len(self.nums) + 1
        self.nums.append((num_id, abstract))
        return num_id

    def numbered(self, items):
        """items: list of (level, text); a fresh numbering instance so the list restarts at 1."""
        num = self.new_num(0)
        for lvl, text in items:
            self.p(text, kind='list', style='ListParagraph', num=num, lvl=lvl)

    def bullets(self, items):
        num = self.new_num(1)
        for lvl, text in items:
            self.p(text, kind='list', style='ListParagraph', num=num, lvl=lvl)

    def list_items(self, n, depth=3):
        pr = self.prose
        items, lvl = [], 0
        for i in range(n):
            if i == 0:
                lvl = 0
            else:
                k = self.r.random()
                if k < 0.35 and lvl < depth - 1:
                    lvl += 1
                elif k < 0.6 and lvl > 0:
                    lvl -= 1
            items.append((lvl, pr.short(1, 2)))
        return items

    # ---- tables ----
    def cell_xml(self, text, width, span=1, vmerge=None, bold=False):
        tcpr = '<w:tcW w:w="%d" w:type="dxa"/>' % width
        if span > 1:
            tcpr += '<w:gridSpan w:val="%d"/>' % span
        if vmerge == 'restart':
            tcpr += '<w:vMerge w:val="restart"/>'
        elif vmerge == 'continue':
            tcpr += '<w:vMerge/>'
        rpr = '<w:b/><w:bCs/>' if bold else ''
        para = self.para_xml(text, spacing='', rpr=rpr) if text else self.para_xml('')
        self.pars.append(('cell', text))
        self.words += len(text.split()) if text else 0
        return '<w:tc><w:tcPr>%s</w:tcPr>%s</w:tc>' % (tcpr, para)

    def table(self, rows, widths):
        """rows: list of rows; a row is a list of cells (text, span, vmerge, bold)."""
        grid = ''.join('<w:gridCol w:w="%d"/>' % w for w in widths)
        xml = ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>'
               '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0"'
               ' w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>%s</w:tblGrid>' % grid)
        for row in rows:
            xml += '<w:tr>'
            col = 0
            for text, span, vmerge, bold in row:
                width = sum(widths[col:col + span])
                xml += self.cell_xml(text, width, span, vmerge, bold)
                col += span
            xml += '</w:tr>'
        xml += '</w:tbl>'
        self.body.append(xml)

    def tbl(self, widths, rows, tblpr=None, tcw=True):
        """A table from explicit parts. rows: dicts {'cells': [...], 'trpr': xml}; a cell is a dict with
        'paras' (a list of (text, para_xml keyword dict)) or 'text', and optionally 'span', 'vmerge', 'fill',
        'valign', 'borders' (tcBorders xml), 'w' (a width overriding the grid), 'wtype'. tcw=False writes no
        cell widths at all, leaving the grid as the only width the file has."""
        grid = ''.join('<w:gridCol w:w="%d"/>' % w for w in widths)
        if tblpr is None:
            tblpr = ('<w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>'
                     '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0"'
                     ' w:noHBand="0" w:noVBand="1"/>')
        xml = '<w:tbl><w:tblPr>%s</w:tblPr><w:tblGrid>%s</w:tblGrid>' % (tblpr, grid)
        for row in rows:
            trpr = row.get('trpr', '')
            xml += '<w:tr>' + ('<w:trPr>%s</w:trPr>' % trpr if trpr else '')
            col = 0
            for cell in row['cells']:
                span = cell.get('span', 1)
                tcpr = ''
                if tcw:
                    width = cell.get('w', sum(widths[col:col + span]))
                    tcpr += '<w:tcW w:w="%d" w:type="%s"/>' % (width, cell.get('wtype', 'dxa'))
                if span > 1:
                    tcpr += '<w:gridSpan w:val="%d"/>' % span
                if cell.get('vmerge') == 'restart':
                    tcpr += '<w:vMerge w:val="restart"/>'
                elif cell.get('vmerge') == 'continue':
                    tcpr += '<w:vMerge/>'
                tcpr += cell.get('borders', '')
                if cell.get('fill'):
                    tcpr += '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % cell['fill']
                if cell.get('valign'):
                    tcpr += '<w:vAlign w:val="%s"/>' % cell['valign']
                paras = cell.get('paras')
                if not paras:                       # a cell with no paragraph at all is a corrupt file to Word
                    paras = [(cell.get('text', ''), {})]
                inner = ''
                for text, kw in paras:
                    inner += self.para_xml(text, **kw)
                    self.pars.append(('cell', text))
                    self.words += len(text.split()) if text else 0
                xml += '<w:tc>%s%s</w:tc>' % (('<w:tcPr>%s</w:tcPr>' % tcpr) if tcpr else '', inner)
                col += span
            xml += '</w:tr>'
        xml += '</w:tbl>'
        self.body.append(xml)

    def merged_table(self):
        """Four columns: a title row spanning all four, a header row, data rows with a vertical merge
        in the first column and a two-column span in one row."""
        pr = self.prose
        widths = [TEXT_W // 4] * 4
        rows = [[(pr.title(2, 4), 4, None, True)],
                [(pr.label(), 1, None, True), (pr.label(), 1, None, True), (pr.label(), 1, None, True), (pr.label(), 1, None, True)]]
        n = self.r.randint(4, 8)
        for i in range(n):
            if i == 1:
                rows.append([(pr.cell(), 1, 'restart', False), (pr.cell(), 1, None, False), (pr.cell(), 1, None, False), (pr.cell(), 1, None, False)])
            elif i == 2:
                rows.append([('', 1, 'continue', False), (pr.cell(), 1, None, False), (pr.cell(), 1, None, False), (pr.cell(), 1, None, False)])
            elif i == 3:
                rows.append([(pr.cell(), 1, None, False), (pr.sentence(6, 14), 2, None, False), (pr.cell(), 1, None, False)])
            else:
                rows.append([(pr.cell(), 1, None, False) for _ in range(4)])
        self.table(rows, widths)

    def plain_table(self, cols, nrows):
        pr = self.prose
        widths = [TEXT_W // cols] * cols
        rows = [[(pr.label(), 1, None, True) for _ in range(cols)]]
        for _ in range(nrows):
            rows.append([(pr.cell(), 1, None, False) for _ in range(cols)])
        self.table(rows, widths)

    def two_column_table(self, nrows):
        pr = self.prose
        widths = [TEXT_W // 3, TEXT_W - TEXT_W // 3]
        rows = [[(pr.label(), 1, None, True), (pr.label(), 1, None, True)]]
        for _ in range(nrows):
            rows.append([(pr.label(), 1, None, False), (pr.short(1, 2), 1, None, False)])
        self.table(rows, widths)

    # ---- pictures and text boxes ----
    def picture(self, kind):
        self.docpr += 1
        fname = 'image%d.png' % (len(self.media) + 1)
        self.media.append((fname, picture_bytes(kind)))
        rid = 'rIdImg%d' % len(self.media)
        cx, cy = 1905000, 1143000
        drawing = (
            '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:docPr id="%d" name="Picture %d"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '<a:graphic><a:graphicData uri="%s"><pic:pic>'
            '<pic:nvPicPr><pic:cNvPr id="%d" name="%s"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
            % (cx, cy, self.docpr, self.docpr, NS_PIC, self.docpr, fname, rid, cx, cy))
        self.body.append(self.para_xml('', jc='center', keep_next=True, runs=drawing))
        self.pars.append(('picture', ''))
        self.p('Figure %d. %s.' % (len(self.media), self.prose.title(3, 6)), kind='caption', jc='center',
               rpr='<w:i/><w:iCs/>')

    def textbox(self, texts):
        self.docpr += 1
        n = self.docpr
        cx, cy = 5486400, 1371600   # 6 in by 1.5 in
        inner = ''.join(self.para_xml(t) for t in texts)
        for t in texts:
            self.pars.append(('textbox', t))
            self.words += len(t.split())
        anchor_id = '%08X' % self.r.getrandbits(31)
        edit_id = '%08X' % self.r.getrandbits(31)
        choice = (
            '<mc:Choice Requires="wps"><w:drawing>'
            '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="%d"'
            ' behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" wp14:anchorId="%s" wp14:editId="%s">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/>'
            '<wp:docPr id="%d" name="Text Box %d"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="%s"><wps:wsp><wps:cNvSpPr txBox="1"/>'
            '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>'
            '<wps:txbx><w:txbxContent>%s</w:txbxContent></wps:txbx>'
            '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"'
            ' anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
            '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>'
            % (251659264 + n, anchor_id, edit_id, cx, cy, n, n, NS_WPS, cx, cy, inner))
        fallback = (
            '<mc:Fallback><w:pict>'
            '<v:rect id="Text Box %d" o:spid="_x0000_s%d" style="position:absolute;margin-left:0;margin-top:0;'
            'width:432pt;height:108pt;z-index:%d;visibility:visible;mso-wrap-style:square;'
            'mso-position-horizontal:absolute;mso-position-horizontal-relative:text;'
            'mso-position-vertical:absolute;mso-position-vertical-relative:text" fillcolor="white" strokeweight=".5pt">'
            '<v:textbox><w:txbxContent>%s</w:txbxContent></v:textbox><w10:wrap type="topAndBottom"/>'
            '</v:rect></w:pict></mc:Fallback>'
            % (n, 1025 + n, 251659264 + n, inner))
        self.body.append(self.p_open() + '<w:r><mc:AlternateContent>%s%s</mc:AlternateContent></w:r></w:p>'
                         % (choice, fallback))
        self.pars.append(('anchor', ''))

    # ---- sections ----
    def sect_xml(self, nextpage=False, borders='', margins=None, continuous=False, header=True, footer=True):
        """borders: a w:pgBorders element; margins: (top, right, bottom, left, header, footer) in twips.
        A section without a header or footer reference shows the previous section's, as Word does."""
        if nextpage and continuous:
            raise SystemExit('%s: a section break is nextPage or continuous, not both' % self.name)
        t, r, b, l, h, f = margins or self.margins
        return ('<w:sectPr>'
                + ('<w:headerReference w:type="default" r:id="rIdHeader1"/>' if header else '')
                + ('<w:footerReference w:type="default" r:id="rIdFooter1"/>' if footer else '')
                + ('<w:type w:val="nextPage"/>' if nextpage else '')
                + ('<w:type w:val="continuous"/>' if continuous else '') +
                '<w:pgSz w:w="%d" w:h="%d"/>'
                '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" w:header="%d" w:footer="%d" w:gutter="0"/>'
                % (PAGE_W, PAGE_H, t, r, b, l, h, f) + borders +
                '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/>'
                '</w:sectPr>')

    # ---- parts ----
    def fonts(self):
        f = self.font
        return '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>' % (f, f, f, f)

    def styles_xml(self):
        f = self.fonts()
        sz = '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (self.sz, self.sz)
        # With a theme part the defaults name the theme's fonts and the Normal style names the real one, which
        # is how Word writes its own files; without one the defaults name the font directly.
        default_fonts = ('<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"'
                         ' w:cstheme="minorBidi"/>' if self.theme else f)
        normal = ('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>'
                  + ('<w:pPr>%s</w:pPr>' % self.normal_ppr if self.normal_ppr else '')
                  + ('<w:rPr>%s</w:rPr>' % f if self.theme else '') + '</w:style>')

        def heading(n, before, size, colour):
            return ('<w:style w:type="paragraph" w:styleId="Heading%d"><w:name w:val="heading %d"/>'
                    '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>'
                    '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="%d" w:after="0"/><w:outlineLvl w:val="%d"/></w:pPr>'
                    '<w:rPr>%s<w:color w:val="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr></w:style>'
                    % (n, n, before, n - 1, f, colour, size, size))

        return XML_HEAD + (
            '<w:styles xmlns:w="%s" xmlns:w14="%s" xmlns:mc="%s" mc:Ignorable="w14">'
            '<w:docDefaults><w:rPrDefault><w:rPr>%s%s<w:lang w:val="en-GB" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '%s'
            '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/>'
            '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
            '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
            '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
            '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
            '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
            '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/>'
            '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
            '%s%s%s'
            '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>'
            '<w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/>'
            '<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/>'
            '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/>'
            '<w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/>'
            '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/>'
            '<w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>'
            '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>'
            '<w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>'
            '%s'
            '</w:styles>'
            % (NS_W, NS_W14, NS_MC, default_fonts, sz, normal,
               heading(1, 240, 32, '1F3864'), heading(2, 40, 26, '2F5496'), heading(3, 40, 24, '1F3763'),
               self.extra_styles))

    def numbering_xml(self):
        f = self.fonts()

        def lvl(i, fmt, text, rpr=''):
            return ('<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="%s"/><w:lvlText w:val="%s"/>'
                    '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="%d" w:hanging="360"/></w:pPr>%s</w:lvl>'
                    % (i, fmt, text, 720 + 720 * i, ('<w:rPr>%s</w:rPr>' % rpr) if rpr else ''))

        decimal = ('<w:abstractNum w:abstractNumId="0"><w:nsid w:val="1A2B3C4D"/><w:multiLevelType w:val="multilevel"/>'
                   '<w:tmpl w:val="2B3C4D5E"/>'
                   + lvl(0, 'decimal', '%1.') + lvl(1, 'decimal', '%1.%2.') + lvl(2, 'decimal', '%1.%2.%3.')
                   + lvl(3, 'lowerLetter', '%4)') + lvl(4, 'lowerRoman', '%5.') + lvl(5, 'decimal', '%6.')
                   + lvl(6, 'lowerLetter', '%7)') + lvl(7, 'lowerRoman', '%8.') + lvl(8, 'decimal', '%9.')
                   + '</w:abstractNum>')
        bullet_fonts = ['<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>',
                        '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/>',
                        '<w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>']
        bullet_chars = ['&#xF0B7;', 'o', '&#xF0A7;']
        bullets = ('<w:abstractNum w:abstractNumId="1"><w:nsid w:val="3C4D5E6F"/><w:multiLevelType w:val="hybridMultilevel"/>'
                   '<w:tmpl w:val="4D5E6F70"/>'
                   + ''.join(lvl(i, 'bullet', bullet_chars[i % 3], bullet_fonts[i % 3]) for i in range(9))
                   + '</w:abstractNum>')
        round_bullets = ''
        if self.round_bullets:
            # Bullets as Word itself writes them: a template code on every level, the font hint, the later levels
            # marked tentative, the first level indented 360 and the rest 720 apart.
            def rlvl(i):
                fonts = ['<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>',
                         '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/>',
                         '<w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>'][i % 3]
                tplc = ['04090001', '04090003', '04090005'][i % 3]
                return ('<w:lvl w:ilvl="%d" w:tplc="%s"%s><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
                        '<w:lvlText w:val="%s"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="%d" w:hanging="360"/></w:pPr>'
                        '<w:rPr>%s%s</w:rPr></w:lvl>'
                        % (i, tplc, ' w:tentative="1"' if i >= 2 else '', bullet_chars[i % 3], 360 + 720 * i, fonts,
                           '<w:color w:val="auto"/>' if i == 0 else ''))
            round_bullets = ('<w:abstractNum w:abstractNumId="2"><w:nsid w:val="4BD81D70"/>'
                             '<w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="5E6F7081"/>'
                             + ''.join(rlvl(i) for i in range(9)) + '</w:abstractNum>')
        nums = ''.join('<w:num w:numId="%d"><w:abstractNumId w:val="%d"/>'
                       '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>' % (n, a)
                       for n, a in self.nums)
        return XML_HEAD + ('<w:numbering xmlns:w="%s" xmlns:w14="%s" xmlns:mc="%s" mc:Ignorable="w14">%s%s%s%s</w:numbering>'
                           % (NS_W, NS_W14, NS_MC, decimal, bullets, round_bullets, nums))

    def theme_xml(self):
        """A theme part of the kind every Word file carries: the Office colours, a major and a minor font."""
        a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        colours = [('dk1', '<a:sysClr val="windowText" lastClr="000000"/>'), ('lt1', '<a:sysClr val="window" lastClr="FFFFFF"/>'),
                   ('dk2', '<a:srgbClr val="44546A"/>'), ('lt2', '<a:srgbClr val="E7E6E6"/>'),
                   ('accent1', '<a:srgbClr val="4472C4"/>'), ('accent2', '<a:srgbClr val="ED7D31"/>'),
                   ('accent3', '<a:srgbClr val="A5A5A5"/>'), ('accent4', '<a:srgbClr val="FFC000"/>'),
                   ('accent5', '<a:srgbClr val="5B9BD5"/>'), ('accent6', '<a:srgbClr val="70AD47"/>'),
                   ('hlink', '<a:srgbClr val="0563C1"/>'), ('folHlink', '<a:srgbClr val="954F72"/>')]
        clr = ''.join('<a:%s>%s</a:%s>' % (n, v, n) for n, v in colours)
        fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        return XML_HEAD + (
            '<a:theme xmlns:a="%s" name="Office Theme"><a:themeElements>'
            '<a:clrScheme name="Office">%s</a:clrScheme>'
            '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
            '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
            '<a:fmtScheme name="Office"><a:fillStyleLst>%s%s%s</a:fillStyleLst>'
            '<a:lnStyleLst><a:ln w="6350">%s</a:ln><a:ln w="12700">%s</a:ln><a:ln w="19050">%s</a:ln></a:lnStyleLst>'
            '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
            '<a:bgFillStyleLst>%s%s%s</a:bgFillStyleLst></a:fmtScheme>'
            '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
            % (a, clr, fill, fill, fill, fill, fill, fill, fill, fill, fill))

    def settings_xml(self):
        return XML_HEAD + (
            '<w:settings xmlns:w="%s" xmlns:w14="%s" xmlns:mc="%s" mc:Ignorable="w14">'
            '<w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '<w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '<w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:uri="http://schemas.microsoft.com/office/word" w:val="0"/>'
            '</w:compat><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/></w:settings>'
            % (NS_W, NS_W14, NS_MC))

    def header_xml(self):
        return XML_HEAD + ('<w:hdr%s>%s</w:hdr>' % (
            ROOT_NS, self.para_xml(self.header_text, style='Header', jc=self.header_align)))

    def footer_xml(self):
        runs = run('Page ') + field('PAGE') + run(' of ') + field('NUMPAGES')
        return XML_HEAD + ('<w:ftr%s>%s</w:ftr>' % (ROOT_NS, self.para_xml('', style='Footer', jc='center', runs=runs)))

    def document_xml(self, final_sect):
        return XML_HEAD + ('<w:document%s><w:body>%s%s</w:body></w:document>' % (ROOT_NS, ''.join(self.body), final_sect))

    def rels_xml(self):
        rels = [('rIdStyles', 'styles', 'styles.xml'), ('rIdNumbering', 'numbering', 'numbering.xml'),
                ('rIdSettings', 'settings', 'settings.xml'), ('rIdHeader1', 'header', 'header1.xml'),
                ('rIdFooter1', 'footer', 'footer1.xml')]
        if self.theme:
            rels.append(('rIdTheme', 'theme', 'theme/theme1.xml'))
        rels += [('rIdImg%d' % (i + 1), 'image', 'media/' + name) for i, (name, _) in enumerate(self.media)]
        return XML_HEAD + ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>'
                           % ''.join('<Relationship Id="%s" Type="%s" Target="%s"/>' % (i, REL_TYPES[t], tg) for i, t, tg in rels))

    def content_types_xml(self):
        wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'
        overrides = [('/word/document.xml', wml + 'document.main+xml'), ('/word/styles.xml', wml + 'styles+xml'),
                     ('/word/numbering.xml', wml + 'numbering+xml'), ('/word/settings.xml', wml + 'settings+xml'),
                     ('/word/header1.xml', wml + 'header+xml'), ('/word/footer1.xml', wml + 'footer+xml'),
                     ('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'),
                     ('/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml')]
        if self.theme:
            overrides.append(('/word/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'))
        return XML_HEAD + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
            '%s</Types>' % ''.join('<Override PartName="%s" ContentType="%s"/>' % o for o in overrides))

    def package_rels_xml(self):
        return XML_HEAD + (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            '</Relationships>')

    def core_xml(self):
        return XML_HEAD + (
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
            ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
            ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>%s</dc:title></cp:coreProperties>' % esc(self.title))

    def app_xml(self):
        return XML_HEAD + (
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
            ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>studio-editor public set</Application></Properties>')

    def write(self, final_sect=None):
        if final_sect is None:
            final_sect = self.final_sect or self.sect_xml()
        parts = [('[Content_Types].xml', self.content_types_xml()), ('_rels/.rels', self.package_rels_xml()),
                 ('word/document.xml', self.document_xml(final_sect)), ('word/_rels/document.xml.rels', self.rels_xml()),
                 ('word/styles.xml', self.styles_xml()), ('word/numbering.xml', self.numbering_xml()),
                 ('word/settings.xml', self.settings_xml()), ('word/header1.xml', self.header_xml()),
                 ('word/footer1.xml', self.footer_xml()), ('docProps/core.xml', self.core_xml()),
                 ('docProps/app.xml', self.app_xml())]
        if self.theme:
            parts.append(('word/theme/theme1.xml', self.theme_xml()))
        path = os.path.join(DOCS, self.name + '.docx')
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            for name, text in parts:
                data = text.encode('ascii')
                ET.fromstring(data)      # every part must parse before it is written
                info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(info, data)
            for name, data in self.media:
                info = zipfile.ZipInfo('word/media/' + name, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(info, data)
        return path


# --------------------------------------------------------------------------------------------------
# The eleven documents
# --------------------------------------------------------------------------------------------------

def body_paragraphs(d, n, lo=2, hi=9):
    for _ in range(n):
        d.p(d.prose.paragraph(lo, hi))


def handbook(rng, pages=60):
    d = Doc('handbook-calibri', 'Harbour Rescue Service Procedures Handbook', 'Calibri', 22, rng)
    pr = d.prose
    d.header_text = 'Procedures Handbook'
    triple = ('<w:pBdr><w:top w:val="triple" w:sz="4" w:space="1" w:color="auto"/>'
              '<w:left w:val="triple" w:sz="4" w:space="4" w:color="auto"/>'
              '<w:bottom w:val="triple" w:sz="4" w:space="1" w:color="auto"/>'
              '<w:right w:val="triple" w:sz="4" w:space="4" w:color="auto"/></w:pBdr>')
    # Title page
    for _ in range(6):
        d.empty()
    d.p('Harbour Rescue Service', kind='title', pbdr=triple, jc='center',
        spacing='<w:spacing w:before="240" w:after="240"/>', rpr='<w:b/><w:bCs/><w:sz w:val="56"/><w:szCs w:val="56"/>')
    d.p('Procedures Handbook', kind='title', jc='center', rpr='<w:sz w:val="36"/><w:szCs w:val="36"/>')
    d.empty()
    d.p('Edition 7, for crews at every station between %s and %s.' % (PLACES[0], PLACES[5]), kind='title', jc='center')
    d.p('Invented text for a layout test; it describes no real service, place or person.', kind='title', jc='center')
    d.page_break()
    # Introduction: plain body pages first, so the edit cases have room
    d.heading(1, 'How to use this handbook')
    body_paragraphs(d, 5, 3, 6)
    d.page_break()
    chapter = 1
    d.heading(1, 'Chapter %d. %s' % (chapter, pr.title(2, 4)))
    d.heading(2, pr.title(2, 5))
    body_paragraphs(d, 4, 3, 6)
    d.page_break()
    for _ in range(3):
        d.heading(2, pr.title(2, 5))
        body_paragraphs(d, 5, 3, 6)
        d.page_break()
    # The remaining pages cycle through the features
    kinds = ['numbered', 'body', 'table', 'body', 'bullets', 'picture', 'body', 'textbox', 'body', 'table',
             'numbered', 'picture', 'body', 'bullets', 'body']
    made = 6
    i = 0
    while made < pages:
        kind = kinds[i % len(kinds)]
        i += 1
        if i % 7 == 1:
            chapter += 1
            d.heading(1, 'Chapter %d. %s' % (chapter, pr.title(2, 4)))
        else:
            d.heading(2 if i % 3 else 3, pr.title(2, 5))
        if kind == 'numbered':
            body_paragraphs(d, 1, 2, 4)
            d.numbered(d.list_items(rng.randint(8, 13)))
            body_paragraphs(d, 1, 2, 4)
        elif kind == 'bullets':
            body_paragraphs(d, 1, 2, 4)
            d.bullets(d.list_items(rng.randint(7, 12)))
            body_paragraphs(d, 1, 2, 4)
        elif kind == 'table':
            body_paragraphs(d, 1, 2, 4)
            d.merged_table()
            body_paragraphs(d, 1, 2, 5)
        elif kind == 'picture':
            body_paragraphs(d, 1, 2, 4)
            d.picture(len(d.media) % 5)
            body_paragraphs(d, 2, 2, 5)
        elif kind == 'textbox':
            body_paragraphs(d, 1, 2, 5)
            d.textbox([pr.short(1, 2), pr.short(1, 2)])
            body_paragraphs(d, 2, 2, 5)
        else:
            body_paragraphs(d, 4, 3, 6)
        made += 1
        if made < pages:
            d.page_break()
    d.p('End of the handbook.')
    return d


def rules(rng, sections=40):
    d = Doc('rules-cambria', 'Harbour Rescue Service Standing Rules', 'Cambria', 22, rng)
    pr = d.prose
    d.header_text = 'Standing Rules'
    d.heading(1, 'Standing Rules of the Harbour Rescue Service')
    for s in range(1, sections + 1):
        if s > 1:
            d.heading(2 if s % 4 else 1, 'Rule %d. %s' % (s, pr.title(2, 5)))
        body_paragraphs(d, 1, 2, 4)
        d.numbered([(0, pr.short(1, 2)) for _ in range(rng.randint(3, 5))])
        if s % 4 == 2:
            d.plain_table(3, rng.randint(3, 5))
        last = pr.paragraph(2, 3)
        if s < sections:
            d.p(last, sect=d.sect_xml(nextpage=True))
        else:
            d.p(last)
    return d


def guide(rng, pages=12):
    d = Doc('guide-arial', 'Field Guide to the Harbour Rescue Service', 'Arial', 20, rng)
    pr = d.prose
    d.header_text = 'Field Guide'
    d.heading(1, 'Field Guide')
    body_paragraphs(d, 2, 3, 5)
    for i in range(pages):
        if i:
            d.page_break()
            d.heading(2, '%d. %s' % (i + 1, pr.title(2, 4)))
        body_paragraphs(d, 1, 2, 4)
        if i in (2, 6, 9):
            d.picture(len(d.media) % 5)
        if i == 4:
            d.textbox([pr.short(1, 2), pr.short(1, 1)])
        if i != 7:
            d.two_column_table(rng.randint(3, 6))
        body_paragraphs(d, 1 if i in (2, 4, 6, 9) else 2, 2, 4)
    return d


def form(rng, blocks=46):
    d = Doc('form-aptos', 'Station Readiness Worksheet', 'Aptos', 22, rng)
    pr = d.prose
    d.header_text = 'Readiness Worksheet'
    d.heading(1, 'Station Readiness Worksheet')
    d.p(pr.paragraph(2, 4))
    for b in range(blocks):
        if b % 6 == 0:
            d.heading(2, 'Part %d. %s' % (b // 6 + 1, pr.title(2, 4)))
        d.p(pr.short(1, 2))
        d.plain_table(rng.randint(2, 5), rng.randint(3, 8))
        d.p(pr.short(1, 3))
    return d


def book(rng, name, title, font, half_points, target_words):
    d = Doc(name, title, font, half_points, rng)
    pr = d.prose
    d.header_text = title
    chapter = 0
    while d.words < target_words:
        chapter += 1
        d.heading(1, 'Chapter %d. %s' % (chapter, pr.title(2, 5)))
        for s in range(rng.randint(3, 5)):
            if s:
                d.heading(2, pr.title(2, 5))
            n = rng.randint(4, 8)
            for k in range(n):
                d.p(pr.paragraph(3, 9))
                if rng.random() < 0.08:
                    d.bullets([(0, pr.short(1, 2)) for _ in range(rng.randint(3, 6))])
        if chapter % 5 == 3:
            d.p(pr.short(1, 2))
            d.plain_table(3, rng.randint(2, 4))
            d.p(pr.short(1, 2))
    return d


# --------------------------------------------------------------------------------------------------
# Three more documents, added 2026-10-29 for the "pages identical" row: the private set loses pages after a
# save on three features the first eight documents did not carry. Each is modelled on the private file that
# loses them, in shape only: a protocol book of bullet lists and boxed tables, a book of notices with page
# borders, and a register of tables with fixed widths.
# --------------------------------------------------------------------------------------------------

def pg_borders(sz, space, colour, offset='page'):
    sides = ''.join('<w:%s w:val="single" w:sz="%d" w:space="%d" w:color="%s"/>' % (s, sz, space, colour)
                    for s in ('top', 'left', 'bottom', 'right'))
    return '<w:pgBorders w:offsetFrom="%s">%s</w:pgBorders>' % (offset, sides)


NO_BORDERS = ('<w:tblBorders>' + ''.join('<w:%s w:val="none" w:sz="0" w:space="0" w:color="auto"/>' % s
                                          for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
              + '</w:tblBorders>')
BLACK_BORDERS = ('<w:tblBorders>' + ''.join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="000000"/>' % s
                                             for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
                 + '</w:tblBorders>')
LOOK = '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
WHITE_BOLD = '<w:b/><w:bCs/><w:color w:val="FFFFFF"/>'


def fixed_table(d, widths, nrows, ind=365):
    """A table of the kind the private operating guidelines carry on most pages: fixed layout, no table
    width, cell widths in twips, no left or right cell margin, row heights with no rule, a merged first
    column and borders switched off cell by cell."""
    pr = d.prose
    tblpr = ('<w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="%d" w:type="dxa"/>%s<w:tblLayout w:type="fixed"/>'
             '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
             '<w:tblLook w:val="01E0" w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1" w:noHBand="0" w:noVBand="0"/>'
             % (ind, BLACK_BORDERS))
    rows = [{'trpr': '<w:trHeight w:val="840"/>',
             'cells': [{'vmerge': 'restart', 'paras': [(pr.label(), {'rpr': '<w:b/><w:bCs/>'})]},
                       {'borders': '<w:tcBorders><w:bottom w:val="nil"/></w:tcBorders>', 'text': pr.sentence(6, 16)}]},
            {'trpr': '<w:trHeight w:val="598"/>',
             'cells': [{'vmerge': 'continue', 'borders': '<w:tcBorders><w:top w:val="nil"/></w:tcBorders>', 'text': ''},
                       {'borders': '<w:tcBorders><w:top w:val="nil"/></w:tcBorders>', 'text': pr.sentence(6, 16)}]}]
    for _ in range(nrows):
        rows.append({'cells': [{'text': pr.label()}, {'text': pr.cell()}]})
    d.tbl(widths, rows, tblpr)


def protocols(rng, count=26):
    """A protocol book in Arial: every protocol its own section, a coloured banner table, bullet lists in
    four spacings (a Bullet style that switches contextual spacing off, plain numbered paragraphs with
    space after, Bullet with space before set to nothing, List Paragraph at an exact line height), a black
    box whose bullets come from the style itself, and a two-column comparison table with a tall row."""
    d = Doc('protocols-arial', 'Harbour Rescue Clinical Protocols', 'Arial', 22, rng)
    pr = d.prose
    d.header_text = 'Clinical Protocols'
    d.theme = True
    d.round_bullets = True
    d.normal_ppr = '<w:spacing w:before="160"/><w:jc w:val="both"/>'
    d.margins = (1440, 1080, 720, 1080, 432, 432)
    box_num = d.new_num(2)
    d.extra_styles = (
        '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Bullet"><w:name w:val="Bullet"/>'
        '<w:basedOn w:val="ListParagraph"/><w:qFormat/><w:pPr><w:spacing w:after="80"/><w:ind w:left="0"/>'
        '<w:contextualSpacing w:val="0"/></w:pPr><w:rPr><w:rFonts w:cs="Arial"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:customStyle="1" w:styleId="BoxBullet"><w:name w:val="BoxBullet"/>'
        '<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:numPr><w:numId w:val="%d"/></w:numPr><w:spacing w:after="80"/>'
        '<w:ind w:left="216" w:hanging="216"/><w:contextualSpacing w:val="0"/><w:jc w:val="left"/></w:pPr>'
        '<w:rPr><w:rFonts w:cs="Arial"/><w:b/><w:color w:val="FFFFFF" w:themeColor="background1"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:customStyle="1" w:styleId="BoxBulletBlack"><w:name w:val="BoxBulletBlack"/>'
        '<w:basedOn w:val="BoxBullet"/><w:qFormat/><w:rPr><w:color w:val="auto"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:customStyle="1" w:styleId="BoxNoBullet"><w:name w:val="BoxNoBullet"/>'
        '<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="80"/><w:jc w:val="left"/></w:pPr>'
        '<w:rPr><w:rFonts w:cs="Arial"/><w:b/><w:color w:val="FFFFFF" w:themeColor="background1"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:customStyle="1" w:styleId="ProtocolHeader"><w:name w:val="ProtocolHeader"/>'
        '<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr>'
        '<w:rPr><w:b/><w:color w:val="000000" w:themeColor="text1"/><w:sz w:val="32"/></w:rPr></w:style>'
        % box_num)
    spacer_ind = '<w:ind w:left="1080" w:hanging="360"/>'
    exact_pre = ('<w:kinsoku w:val="0"/><w:overflowPunct w:val="0"/><w:autoSpaceDE w:val="0"/>'
                 '<w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/>')

    def items(n):
        out = []
        for i in range(n):
            k = rng.random()
            lvl = 0 if i == 0 or k < 0.55 else 2
            text = pr.label() if rng.random() < 0.15 else pr.sentence(4, 30)
            out.append((lvl, text))
        return out

    def bullet_block(variant, n):
        num = d.new_num(2)
        for lvl, text in items(n):
            if variant == 0:
                d.p(text, kind='list', style='Bullet', num=num, lvl=lvl)
            elif variant == 1:
                d.p(text, kind='list', num=num, lvl=lvl, spacing='<w:spacing w:after="80"/>')
            elif variant == 2:
                d.p(text, kind='list', style='Bullet', num=num, lvl=min(lvl, 1), spacing='<w:spacing w:before="0"/>')
            else:
                d.p(text, kind='list', style='ListParagraph', num=num, lvl=0, pre=exact_pre,
                    spacing='<w:spacing w:before="0" w:after="0" w:line="221" w:lineRule="exact"/>',
                    ind='<w:ind w:right="73"/>', jc='left')

    def subheading(text):
        d.p(text, kind='heading', keep_next=True, spacing='<w:spacing w:after="120"/>', rpr='<w:b/><w:bCs/>')

    def banner(p):
        d.tbl([4485, 4505], [
            {'trpr': '<w:trHeight w:val="989"/><w:jc w:val="center"/>',
             'cells': [{'fill': 'FF0000', 'valign': 'center', 'paras': [('Protocol %d' % p, {'style': 'ProtocolHeader'})]},
                       {'fill': '1F4E79', 'valign': 'center',
                        'paras': [(pr.title(2, 3), {'style': 'ProtocolHeader', 'rpr': WHITE_BOLD + '<w:sz w:val="32"/>'})]}]},
            {'trpr': '<w:jc w:val="center"/>',
             'cells': [{'text': pr.short(1, 1)}, {'text': pr.short(1, 1)}]}])

    def black_box():
        tblpr = '<w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>%s%s' % (NO_BORDERS, LOOK)
        rows = [{'cells': [{'fill': '000000', 'paras': [(pr.title(2, 4), {'style': 'BoxNoBullet'})]
                            + [(pr.sentence(4, 18), {'style': 'BoxBullet'}) for _ in range(rng.randint(2, 4))]}]}]
        for _ in range(rng.randint(2, 5)):
            rows.append({'cells': [{'paras': [(pr.sentence(4, 22), {'style': 'BoxBulletBlack'})
                                              for _ in range(rng.randint(1, 3))]}]})
        d.tbl([10070], rows, tblpr)

    def comparison():
        num_a, num_b = d.new_num(2), d.new_num(2)
        left = [(t, {'style': 'Bullet', 'num': num_a, 'lvl': l}) for l, t in items(rng.randint(4, 7))]
        right = [(t, {'style': 'Bullet', 'num': num_b, 'lvl': l}) for l, t in items(rng.randint(4, 7))]
        d.tbl([5125, 5035], [
            {'trpr': '<w:trHeight w:val="872"/>',
             'cells': [{'fill': 'FF0000', 'valign': 'center', 'paras': [(pr.label(), {'style': 'ProtocolHeader'})]},
                       {'fill': '1F4E79', 'valign': 'center',
                        'paras': [(pr.label(), {'style': 'ProtocolHeader', 'rpr': WHITE_BOLD + '<w:sz w:val="32"/>'})]}]},
            {'trpr': '<w:trHeight w:val="5102"/>', 'cells': [{'paras': left}, {'paras': right}]}])

    d.heading(1, 'Clinical Protocols of the Harbour Rescue Service')
    d.p(pr.paragraph(2, 4))
    d.p('Invented text for a layout test; it describes no real service, place, person or treatment.')
    d.p('', kind='empty', style='Bullet', ind=spacer_ind, sect=d.sect_xml())
    for p in range(1, count + 1):
        d.heading(1, 'Protocol %d. %s' % (p, pr.title(2, 4)))
        banner(p)
        subheading(pr.label())
        d.p(pr.paragraph(1, 3))
        bullet_block(p % 4, rng.randint(7, 13))
        if p % 3 != 1:
            subheading(pr.label())
            bullet_block((p + 1) % 4, rng.randint(5, 9))
        if p % 2 == 0:
            black_box()
        else:
            comparison()
        d.p('', kind='empty', style='Bullet', ind=spacer_ind)
        if p % 4 == 3:
            subheading(pr.label())
            bullet_block(0, rng.randint(4, 8))
        if p < count:
            d.p('', kind='empty', style='Bullet', ind=spacer_ind, sect=d.sect_xml())
    d.p('End of the protocols.')
    return d


def notices(rng, sections=18):
    """Standing notices in Calibri: sections with page borders measured from the page edge, in two weights,
    beside sections with none and one with a border measured from the text; a continuous first section;
    margins that change from section to section; the fixed tables of the guidelines on most pages."""
    d = Doc('notices-calibri', 'Harbour Rescue Service Standing Notices', 'Calibri', 22, rng)
    pr = d.prose
    d.header_text = 'Standing Notices'
    d.theme = True
    d.round_bullets = True
    heavy = pg_borders(48, 24, '5F5F5F')
    black = pg_borders(36, 28, '000000')
    thin = pg_borders(12, 4, '1F3864', offset='text')
    # (borders, margins, header, footer, pages of content)
    kinds = [(heavy, (660, 720, 280, 720, 720, 720), True, True, 1),
             (black, (1820, 720, 280, 720, 0, 0), False, False, 1),
             ('', (640, 720, 1340, 720, 0, 1148), False, True, 2),
             (heavy, (660, 720, 280, 720, 720, 720), True, True, 3),
             (thin, (1440, 1440, 1440, 1440, 720, 720), True, True, 1),
             (black, (1820, 720, 280, 720, 0, 0), False, False, 2)]

    def bullets(n):
        num = d.new_num(2)
        for _ in range(n):
            d.p(pr.sentence(4, 24), kind='list', style='ListParagraph', num=num, lvl=0)

    # Title section: continuous, with the heavy border
    for _ in range(5):
        d.empty()
    d.p('Harbour Rescue Service', kind='title', jc='center', rpr='<w:b/><w:bCs/><w:sz w:val="56"/><w:szCs w:val="56"/>')
    d.p('Standing Notices', kind='title', jc='center', rpr='<w:sz w:val="36"/><w:szCs w:val="36"/>')
    d.empty()
    d.p('Invented text for a layout test; it describes no real service, place or person.', kind='title', jc='center')
    d.empty()
    fixed_table(d, [5040, 5040], 2)
    d.p('', kind='empty', sect=d.sect_xml(borders=heavy, margins=(660, 720, 280, 720, 720, 720), continuous=True))
    for s in range(1, sections + 1):
        borders, margins, header, footer, pages = kinds[s % len(kinds)]
        d.heading(1, 'Notice %d. %s' % (s, pr.title(2, 5)))
        for page in range(pages):
            if page:
                d.heading(2, pr.title(2, 4))
            d.p(pr.paragraph(2, 4))
            if (s + page) % 2:
                fixed_table(d, [5040, 5040], rng.randint(1, 3))
                d.p(pr.paragraph(1, 2))
            bullets(rng.randint(3, 6))
            d.p(pr.paragraph(1, 3))
            if pages > 1 and page < pages - 1:
                d.page_break()
        last = pr.short(1, 2)
        if s < sections:
            d.p(last, sect=d.sect_xml(borders=borders, margins=margins, header=header, footer=footer))
        else:
            d.p(last)
    borders, margins, header, footer, _ = kinds[sections % len(kinds)]
    d.final_sect = d.sect_xml(borders=borders, margins=margins, header=header, footer=footer)
    return d


def register(rng, blocks=42):
    """An equipment register in Aptos: tables of seven kinds between short paragraphs. Widths from the grid
    alone, widths in twips with a fixed layout, a table in percent with cell spacing and centred rows, rows
    with a height and centred rows, a floating table, cells with automatic widths, and merged cells with
    uneven fixed widths."""
    d = Doc('register-aptos', 'Station Equipment Register', 'Aptos', 22, rng)
    pr = d.prose
    d.header_text = 'Equipment Register'
    d.heading(1, 'Station Equipment Register')
    d.p(pr.paragraph(2, 4))

    def header_cells(n):
        return [{'paras': [(pr.label(), {'rpr': '<w:b/><w:bCs/>'})]} for _ in range(n)]

    def plain_rows(cols, n):
        return [{'cells': [{'text': pr.cell()} for _ in range(cols)]} for _ in range(n)]

    def grid_only():
        tblpr = '<w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/>' + LOOK
        d.tbl([2000, 4000, 3360], [{'cells': header_cells(3)}] + plain_rows(3, rng.randint(3, 7)), tblpr, tcw=False)

    def percent():
        tblpr = ('<w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/><w:tblCellSpacing w:w="15" w:type="dxa"/>'
                 '<w:tblBorders><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
                 '<w:tblCellMar><w:top w:w="15" w:type="dxa"/><w:left w:w="15" w:type="dxa"/>'
                 '<w:bottom w:w="15" w:type="dxa"/><w:right w:w="15" w:type="dxa"/></w:tblCellMar>' + LOOK)
        spacing = '<w:tblCellSpacing w:w="15" w:type="dxa"/><w:jc w:val="center"/>'
        rows = [{'trpr': '<w:tblHeader/>' + spacing, 'cells': header_cells(4)}]
        for _ in range(rng.randint(3, 6)):
            rows.append({'trpr': spacing, 'cells': [{'text': pr.cell()} for _ in range(4)]})
        for row in rows:
            for cell in row['cells']:
                cell['w'], cell['wtype'] = 1250, 'pct'
        d.tbl([2340, 2340, 2340, 2340], rows, tblpr)

    def centred_rows():
        rows = [{'trpr': '<w:trHeight w:val="567"/><w:jc w:val="center"/>',
                 'cells': [dict(c, fill='D9E2F3', valign='center') for c in header_cells(2)]}]
        for _ in range(rng.randint(3, 6)):
            rows.append({'trpr': '<w:trHeight w:val="567"/><w:jc w:val="center"/>',
                         'cells': [{'text': pr.label(), 'valign': 'center'}, {'text': pr.cell(), 'valign': 'center'}]})
        d.tbl([3000, 3000], rows)

    def floating():
        tblpr = ('<w:tblStyle w:val="TableGrid"/><w:tblpPr w:leftFromText="180" w:rightFromText="180"'
                 ' w:vertAnchor="text" w:tblpY="1"/><w:tblOverlap w:val="never"/><w:tblW w:w="0" w:type="auto"/>' + LOOK)
        d.tbl([3600, 3600], [{'cells': header_cells(2)}] + plain_rows(2, rng.randint(2, 4)), tblpr)
        d.p(pr.paragraph(2, 3))

    def auto_cells():
        rows = [{'cells': [dict(c, w=0, wtype='auto') for c in header_cells(3)]}]
        for _ in range(rng.randint(3, 6)):
            rows.append({'cells': [{'text': pr.cell(), 'w': 0, 'wtype': 'auto'} for _ in range(3)]})
        d.tbl([3120, 3120, 3120], rows)

    def merged_fixed():
        tblpr = '<w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/>' + LOOK
        rows = [{'cells': [{'span': 4, 'paras': [(pr.title(2, 4), {'rpr': '<w:b/><w:bCs/>'})]}]},
                {'cells': header_cells(4)},
                {'cells': [{'vmerge': 'restart', 'text': pr.cell()}, {'text': pr.cell()}, {'text': pr.cell()}, {'text': pr.cell()}]},
                {'cells': [{'vmerge': 'continue', 'text': ''}, {'text': pr.cell()}, {'span': 2, 'text': pr.sentence(6, 14)}]}]
        for _ in range(rng.randint(1, 4)):
            rows.append({'cells': [{'text': pr.cell()} for _ in range(4)]})
        d.tbl([1500, 2500, 2500, 2860], rows, tblpr)

    kinds = [grid_only, lambda: fixed_table(d, [4500, 4860], rng.randint(1, 4), ind=0), percent, centred_rows,
             floating, auto_cells, merged_fixed]
    for b in range(blocks):
        if b % 7 == 0:
            d.heading(2, 'Part %d. %s' % (b // 7 + 1, pr.title(2, 4)))
        d.p(pr.short(1, 2))
        kinds[b % len(kinds)]()
        d.p(pr.short(1, 3))
        if b % 3 == 2 and b < blocks - 1:
            d.page_break()
    return d


# --------------------------------------------------------------------------------------------------
# The edit cases: phrases that occur exactly once in the handbook
# --------------------------------------------------------------------------------------------------

PUNCT = re.compile(r'[,;:()]')


def choose_phrases(pars):
    """pars: (kind, text) in document order. Returns the set.json 'cases' block, or raises."""
    full = '\n'.join(t for _, t in pars)

    def unique(phrase):
        return full.count(phrase) == 1

    def pick(words, where, allow_punct):
        # 4..8 words; without inner punctuation unless allow_punct
        if True:
            for n in range(4, 9):
                if n > len(words):
                    break
                if where == 'first':
                    starts = [0]
                elif where == 'last':
                    starts = [len(words) - n]
                else:
                    starts = [s for s in range(2, len(words) - n) if s + n <= len(words) - 1]
                for s in starts:
                    chunk = words[s:s + n]
                    phrase = ' '.join(chunk)
                    inner = phrase[:-1] if where == 'last' else phrase
                    if where != 'last' and phrase.endswith('.'):
                        continue
                    if not allow_punct and PUNCT.search(inner):
                        continue
                    if unique(phrase):
                        return phrase
        return None

    limit = min(120, len(pars))
    body_idx = [i for i in range(limit) if pars[i][0] == 'body' and len(pars[i][1].split(' ')) >= 12]
    chosen = []   # indices already used

    def free(i):
        return all(abs(i - j) >= 2 for j in chosen)

    def take(where):
        # a clean phrase in any paragraph beats a punctuated one in an earlier paragraph
        for allow_punct in (False, True):
            for i in body_idx:
                if i in chosen or not free(i):
                    continue
                phrase = pick(pars[i][1].split(' '), where, allow_punct)
                if phrase:
                    chosen.append(i)
                    return i, phrase
        raise SystemExit('no paragraph left for an edit case (%s)' % where)

    cases = {}
    i_type, p_type = take('middle')
    i_split, p_split = take('first')
    i_del, p_del = take('middle')
    i_copy, p_copy = take('first')
    i_before1, p_before1 = take('first')
    i_cut, p_cut = take('first')
    i_before2, p_before2 = take('first')
    i_delp, p_delp = take('first')
    i_newp, p_newp = take('last')
    # the join: two plain body paragraphs in a row, clear of everything chosen so far
    join = None
    for i in body_idx:
        if i + 1 in body_idx and free(i) and free(i + 1):
            a = pick(pars[i][1].split(' '), 'first', True)
            b = pick(pars[i + 1][1].split(' '), 'first', True)
            if a and b:
                join = (i, a, i + 1, b)
                chosen.extend([i, i + 1])
                break
    if join is None:
        raise SystemExit('no two adjacent body paragraphs for the join case')
    cases = {
        'type': {'in': p_type, 'mark': 'EDITMARKONE'},
        'split': {'in': p_split, 'mark': 'SPLITMARK'},
        'delete_text': {'in': p_del},
        'copy_paste': {'copy': p_copy, 'before': p_before1},
        'move': {'cut': p_cut, 'before': p_before2},
        'delete_paragraph': {'in': p_delp},
        'new_paragraph': {'after': p_newp, 'mark': 'NEWPARAMARK'},
        'join': {'first': join[1], 'second': join[3]},
    }
    # ---- assert every condition on the generated text ----
    phrases = [p_type, p_split, p_del, p_copy, p_before1, p_cut, p_before2, p_delp, p_newp, join[1], join[3]]
    idx = [i_type, i_split, i_del, i_copy, i_before1, i_cut, i_before2, i_delp, i_newp, join[0], join[2]]
    assert len(set(phrases)) == 11, 'phrases are not all different'
    assert len(set(idx)) == 11, 'paragraphs are not all different'
    for ph in phrases:
        assert 4 <= len(ph.split(' ')) <= 8, 'phrase length: ' + ph
        assert full.count(ph) == 1, 'phrase not unique: ' + ph
    for i in idx:
        assert i < 120 and pars[i][0] == 'body', 'not a plain body paragraph in the first 120: %d' % i
    for a in idx:
        for b in idx:
            if a != b and abs(a - b) < 2 and {a, b} != {join[0], join[2]}:
                raise SystemExit('paragraphs %d and %d are neighbours' % (a, b))
    assert join[2] == join[0] + 1, 'join paragraphs are not adjacent'

    def text_of(i):
        return pars[i][1]

    def words_of(i):
        return text_of(i).split(' ')

    for i, ph in ((i_split, p_split), (i_copy, p_copy), (i_before1, p_before1), (i_cut, p_cut),
                  (i_before2, p_before2), (i_delp, p_delp), (join[0], join[1]), (join[2], join[3])):
        assert text_of(i).startswith(ph + ' '), 'not the first words: ' + ph
    for i, ph in ((i_type, p_type), (i_del, p_del)):
        t = text_of(i)
        pos = t.find(ph)
        assert pos > 0 and t[pos - 1] == ' ', 'not in the middle: ' + ph
        rest = t[pos + len(ph):]
        assert rest.startswith(' ') and len(rest.split()) >= 1, 'phrase is at the end of its paragraph: ' + ph
    assert text_of(i_newp).endswith(' ' + p_newp) and p_newp.endswith('.'), 'not the last words: ' + p_newp
    return cases


# --------------------------------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------------------------------

def write_text(path, text):
    data = text.encode('ascii')
    for b in data:
        if b != 0x0A and not (0x20 <= b <= 0x7E):
            raise SystemExit('non-ASCII or control byte in %s' % path)
    with open(path, 'wb') as f:
        f.write(data)


def scan_file(path):
    with open(path, 'rb') as f:
        data = f.read()
    bad = [b for b in data if b != 0x0A and not (0x20 <= b <= 0x7E)]
    return len(bad)


def main():
    os.makedirs(DOCS, exist_ok=True)
    master = random.Random(SEED)
    seeds = [master.getrandbits(32) for _ in range(11)]   # the first eight are the 2026-10-29 documents, unchanged
    vocab = set(NOUNS) | set(ADJECTIVES) | set(VERBS) | set(ADVERBS) | set(FUNCTION_WORDS) | set(UNITS)
    print('vocabulary: %d distinct words' % len(vocab))

    docs = []
    hb = handbook(random.Random(seeds[0]))
    cases = choose_phrases(hb.pars)
    docs.append(hb)
    docs.append(rules(random.Random(seeds[1])))
    docs.append(guide(random.Random(seeds[2])))
    docs.append(form(random.Random(seeds[3])))
    docs.append(book(random.Random(seeds[4]), 'book-calibri', 'Tides at Greyhaven', 'Calibri', 22, 236000))
    docs.append(book(random.Random(seeds[5]), 'book-cambria', 'The Lantern Quay Years', 'Cambria', 22, 33000))
    docs.append(book(random.Random(seeds[6]), 'book-arial', 'Nine Winters on Kestrel Bank', 'Arial', 21, 36000))
    docs.append(book(random.Random(seeds[7]), 'book-aptos', 'A Season at Saltmarsh Point', 'Aptos', 22, 33000))
    docs.append(protocols(random.Random(seeds[8])))
    docs.append(notices(random.Random(seeds[9])))
    docs.append(register(random.Random(seeds[10])))
    for d in docs:
        path = d.write()
        print('%-22s %9d bytes  %7d words  %5d paragraphs' % (d.name + '.docx', os.path.getsize(path), d.words, len(d.pars)))

    set_json = {
        'name': 'public',
        'documents': [d.name for d in docs],
        'flow': ['handbook-calibri', 'rules-cambria'],
        'edit': {'file': 'handbook-calibri', 'cases': cases},
    }
    write_text(os.path.join(HERE, 'set.json'), json.dumps(set_json, indent=2) + '\n')
    for k, v in cases.items():
        print('  %-17s %s' % (k, json.dumps(v)))
    for name in ('make.py', 'set.json', 'README.md'):
        path = os.path.join(HERE, name)
        if os.path.exists(path):
            print('byte scan %-10s %s' % (name, 'clean' if scan_file(path) == 0 else 'HAS BAD BYTES'))


if __name__ == '__main__':
    main()

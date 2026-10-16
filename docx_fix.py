#!/usr/bin/env python3
"""Check and fix a .docx for digital distribution, as specified in spec.md.

Edits the OOXML parts directly. Every rule counts its changes; full passes repeat until a
pass makes 0 changes (convergence). Destructive orphan cleanup is verified and reverted on
failure. The input file is never overwritten.

Writes:
    <out>.docx          the fixed document (fix mode only)
    <out>.report.md     per-pass change counts, findings, items left for human review

Usage:
    uv run docx_fix.py check <in.docx>                report what would change; exit 1 if anything
    uv run docx_fix.py fix <in.docx> [<out.docx>]     default out: <in>_fixed.docx

Options:
    --heading-style ID=N   treat paragraph style ID as heading level N (repeatable), for
                           custom heading styles that resolve to no outline level
    --leadin ID            accept a reviewed '.'-ending lead-in (ID from the report) and convert
                           its final '.' to ':' (repeatable)
    --strip-rsid           also strip rsid editing fingerprints, set removePersonalInformation
    --only N[,N...]        run only the rules with these spec numbers (e.g. --only 8.1,8.2 accepts
                           tracked changes and removes comments, leaving layout and text alone)
    --font NAME            set the theme heading and body fonts to NAME (with rule 11.6)
    --scheme FILE.toml     colour scheme applied after the house style (rule 11.6): [text] style
                           (id or name) = RRGGBB|auto; [table] header_fill = RRGGBB

See spec.md.
"""
import argparse
import copy
import hashlib
import os
import posixpath
import re
import sys
import tomllib
import zipfile
from collections import Counter, defaultdict
from urllib.parse import unquote

from lxml import etree

MAX_PASSES = 10

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'o': 'urn:schemas-microsoft-com:office:office',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'pr': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
}
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
PARSER = etree.XMLParser(huge_tree=True, remove_blank_text=False)
CONTENT_TYPES = '[Content_Types].xml'


def q(tag):
    """'w:p' -> '{namespace}p'."""
    prefix, local = tag.split(':')
    return '{%s}%s' % (NS[prefix], local)


P, R, T, TBL, TR, PPR, RPR, SDT = (q('w:' + t) for t in ('p', 'r', 't', 'tbl', 'tr', 'pPr', 'rPr', 'sdt'))


def wv(el, attr='val'):
    return None if el is None else el.get(q('w:' + attr))


def local(el):
    return etree.QName(el).localname if isinstance(el.tag, str) else None


def attrs(el):
    return {etree.QName(k).localname: v for k, v in el.attrib.items()}


def mk(tag, values=None):
    el = etree.Element(q(tag))
    for k, v in (values or {}).items():
        el.set(q('w:' + k), v)
    return el


def is_on(el):
    """On/off property: None when absent, else its boolean value."""
    if el is None:
        return None
    return wv(el) in (None, '1', 'true', 'on')


def remove(el):
    el.getparent().remove(el)


def unwrap(el):
    parent = el.getparent()
    idx = parent.index(el)
    for i, child in enumerate(list(el)):
        parent.insert(idx + i, child)
    parent.remove(el)


# ---------------------------------------------------------------------------------------------
# Schema order (Appendix B): children of these elements must follow the ECMA-376 CT_* sequence.

PPR_ORDER = (
    'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers '
    'pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE '
    'autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents '
    'suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle '
    'rPr sectPr pPrChange').split()
TRPR_ORDER = ('cnfStyle divId gridBefore gridAfter wBefore wAfter cantSplit trHeight tblHeader '
              'tblCellSpacing jc hidden ins del trPrChange').split()
TBLPR_ORDER = ('tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize '
               'tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook '
               'tblCaption tblDescription tblPrChange').split()
SECTPR_ORDER = ('headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc '
                'pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg '
                'textDirection bidi rtlGutter docGrid printerSettings sectPrChange').split()
SETTINGS_ORDER = (
    'writeProtection view zoom removePersonalInformation removeDateAndTime '
    'doNotDisplayPageBoundaries displayBackgroundShape printPostScriptOverText '
    'printFractionalCharacterWidth printFormsData embedTrueTypeFonts embedSystemFonts '
    'saveSubsetFonts saveFormsData mirrorMargins alignBordersAndEdges '
    'bordersDoNotSurroundHeader bordersDoNotSurroundFooter gutterAtTop hideSpellingErrors '
    'hideGrammaticalErrors activeWritingStyle proofState formsDesign attachedTemplate linkStyles '
    'stylePaneFormatFilter stylePaneSortMethod documentType mailMerge revisionView '
    'trackRevisions doNotTrackMoves doNotTrackFormatting documentProtection autoFormatOverride '
    'styleLockTheme styleLockQFSet defaultTabStop autoHyphenation consecutiveHyphenLimit '
    'hyphenationZone doNotHyphenateCaps showEnvelope summaryLength clickAndTypeStyle '
    'defaultTableStyle evenAndOddHeaders bookFoldRevPrinting bookFoldPrinting '
    'bookFoldPrintingSheets drawingGridHorizontalSpacing drawingGridVerticalSpacing '
    'displayHorizontalDrawingGridEvery displayVerticalDrawingGridEvery '
    'doNotUseMarginsForDrawingGridOrigin drawingGridHorizontalOrigin drawingGridVerticalOrigin '
    'doNotShadeFormData noPunctuationKerning characterSpacingControl printTwoOnOne '
    'strictFirstAndLastChars noLineBreaksAfter noLineBreaksBefore savePreviewPicture '
    'doNotValidateAgainstSchema saveInvalidXml ignoreMixedContent alwaysShowPlaceholderText '
    'doNotDemarcateInvalidXml saveXmlDataOnly useXSLTWhenSaving saveThroughXslt showXMLTags '
    'alwaysMergeEmptyNamespace updateFields hdrShapeDefaults footnotePr endnotePr compat docVars '
    'rsids mathPr attachedSchema themeFontLang clrSchemeMapping doNotIncludeSubdocsInStats '
    'doNotAutoCompressPictures forceUpgrade captions readModeInkLockDown smartTagType '
    'schemaLibrary shapeDefaults doNotEmbedSmartTags decimalSymbol listSeparator').split()
ORDERS = {'pPr': PPR_ORDER, 'trPr': TRPR_ORDER, 'tblPr': TBLPR_ORDER, 'sectPr': SECTPR_ORDER,
          'settings': SETTINGS_ORDER}


def insert_ordered(parent, child, order):
    """Insert before the first w: child whose order index is greater (never plain append)."""
    idx = order.index(local(child))
    for i, c in enumerate(parent):
        name = local(c)
        if name in order and etree.QName(c).namespace == NS['w'] and order.index(name) > idx:
            parent.insert(i, child)
            return child
    parent.append(child)
    return child


def ensure(parent, name, order):
    """Return (element, created). Tests `is None`, never truthiness (Appendix A)."""
    el = parent.find(q('w:' + name))
    if el is not None:
        return el, False
    return insert_ordered(parent, mk('w:' + name), order), True


def set_flag(parent, name, order):
    """Turn an on/off property on. Returns 1 if anything changed."""
    el, created = ensure(parent, name, order)
    if created:
        return 1
    if not is_on(el):
        el.attrib.pop(q('w:val'))
        return 1
    return 0


def set_attrs(el, values):
    if attrs(el) == values:
        return 0
    el.attrib.clear()
    for k, v in values.items():
        el.set(q('w:' + k), v)
    return 1


def ppr(p):
    el = p.find(PPR)
    if el is None:
        el = mk('w:pPr')
        p.insert(0, el)
    return el


# ---------------------------------------------------------------------------------------------
# Package: the ZIP of XML parts.

def rels_path(part):
    d, b = posixpath.split(part)
    return posixpath.join(d, '_rels', b + '.rels')


class Package:
    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.names = [i.filename for i in z.infolist() if not i.is_dir()]
            self.raw = {n: z.read(n) for n in self.names}
        self.trees = {}

    def xml(self, name):
        if name is None or name not in self.raw:
            return None
        if name not in self.trees:
            self.trees[name] = etree.fromstring(self.raw[name], PARSER)
        return self.trees[name]

    def rel_target(self, rels_name, rel):
        if rel.get('TargetMode') == 'External':
            return None
        target = unquote(rel.get('Target', ''))
        if target.startswith('/'):
            return target[1:]
        src_dir = posixpath.dirname(posixpath.dirname(rels_name))
        return posixpath.normpath(posixpath.join(src_dir, target))

    def rels(self, part):
        """(rel element, resolved target, type) for each relationship of `part`."""
        root = self.xml(rels_path(part))
        if root is None:
            return []
        name = rels_path(part)
        return [(rel, self.rel_target(name, rel), rel.get('Type', '')) for rel in root]

    def related(self, part, type_suffix):
        return [t for _, t, ty in self.rels(part) if ty.endswith('/' + type_suffix) and t in self.raw]

    def delete(self, name):
        if name not in self.raw:
            return False
        self.names.remove(name)
        del self.raw[name]
        self.trees.pop(name, None)
        ct = self.xml(CONTENT_TYPES)
        for o in ct.findall(q('ct:Override')):
            if o.get('PartName') == '/' + name:
                ct.remove(o)
        for rels_name in [n for n in self.names if n.endswith('.rels')]:
            for rel in list(self.xml(rels_name)):
                if self.rel_target(rels_name, rel) == name:
                    remove(rel)
        self.delete(rels_path(name))
        return True

    def serialize(self):
        for name, tree in self.trees.items():
            self.raw[name] = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)

    def save(self, path):
        """Appendix E: [Content_Types].xml first and stored, the rest deflated, no dir entries."""
        self.serialize()
        tmp = path + '.part'
        with zipfile.ZipFile(tmp, 'w') as z:
            z.writestr(zipfile.ZipInfo(CONTENT_TYPES, (1980, 1, 1, 0, 0, 0)), self.raw[CONTENT_TYPES],
                       compress_type=zipfile.ZIP_STORED)
            for name in self.names:
                if name != CONTENT_TYPES:
                    z.writestr(zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0)), self.raw[name],
                               compress_type=zipfile.ZIP_DEFLATED)
        os.replace(tmp, path)


# ---------------------------------------------------------------------------------------------
# Document model helpers.

class Styles:
    def __init__(self, root):
        self.by_id, self.by_name = {}, {}
        self.default_p = None
        self.doc_bold = False
        if root is None:
            return
        for s in root.findall(q('w:style')):
            sid = s.get(q('w:styleId'))
            self.by_id[sid] = s
            name = s.find(q('w:name'))
            if name is not None:
                self.by_name[wv(name)] = sid
            if s.get(q('w:type')) == 'paragraph' and s.get(q('w:default')) in ('1', 'true'):
                self.default_p = sid
        self.doc_bold = bool(is_on(root.find('w:docDefaults/w:rPrDefault/w:rPr/w:b', NS)))

    def chain(self, sid):
        seen = set()
        while sid and sid not in seen and sid in self.by_id:
            seen.add(sid)
            s = self.by_id[sid]
            yield s
            sid = wv(s.find(q('w:basedOn')))

    def outline(self, sid):
        for s in self.chain(sid):
            lvl = s.find('w:pPr/w:outlineLvl', NS)
            if lvl is not None:
                return int(wv(lvl))
        return None

    def bold(self, sid):
        for s in self.chain(sid):
            v = is_on(s.find('w:rPr/w:b', NS))
            if v is not None:
                return v
        return None


class Doc:
    def __init__(self, pkg, opts):
        self.pkg, self.opts = pkg, opts
        self.main = pkg.related('', 'officeDocument')[0]
        self.findings = defaultdict(list)   # review items, rebuilt every pass
        self.log = defaultdict(list)        # what was changed, kept across passes
        self.stats = {}

    def part(self, rtype):
        found = self.pkg.related(self.main, rtype)
        return found[0] if found else None

    def xml(self, name):
        return self.pkg.xml(name)

    @property
    def stories(self):
        parts = [self.main]
        for rtype in ('header', 'footer', 'footnotes', 'endnotes'):
            parts += self.pkg.related(self.main, rtype)
        return [self.xml(p) for p in parts]

    @property
    def word_parts(self):
        """Every XML part under the main document's folder (styles, numbering, settings, ...)."""
        folder = posixpath.dirname(self.main)
        return [self.xml(n) for n in self.pkg.names
                if n.startswith(folder + '/') and n.endswith('.xml') and '/_rels/' not in n]

    @property
    def settings(self):
        return self.xml(self.part('settings'))

    def styles(self):
        return Styles(self.xml(self.part('styles')))

    def heading_level(self, p, st):
        """1-based heading level, resolving direct outlineLvl, custom styles via basedOn, overrides."""
        pp = p.find(PPR)
        lvl = wv(pp.find(q('w:outlineLvl'))) if pp is not None else None
        if lvl is not None:
            return int(lvl) + 1 if int(lvl) < 9 else None
        sid = p_style(p, st)
        if sid in self.opts.heading_styles:
            return self.opts.heading_styles[sid]
        lvl = st.outline(sid)
        return lvl + 1 if lvl is not None and lvl < 9 else None


def p_style(p, st):
    return wv(p.find('w:pPr/w:pStyle', NS)) or st.default_p


def is_list(p):
    num = p.find('w:pPr/w:numPr', NS)
    return num is not None and wv(num.find(q('w:numId'))) != '0'


CONTAINER_TAGS = [q('w:' + t) for t in ('body', 'tc', 'txbxContent', 'hdr', 'ftr', 'footnote', 'endnote',
                                        'sdtContent')]
BLOCK_TAGS = {P, TBL, SDT, q('w:altChunk'), q('w:customXml')}


def containers(root):
    return list(root.iter(*CONTAINER_TAGS))


def blocks(container):
    """Block-level children; bookmarks, proofing and range markers between them are transparent."""
    return [c for c in container if c.tag in BLOCK_TAGS]


# Paragraph text as a flat string with an offset map back to its w:t elements (Appendix C).
# Non-text content becomes a \x00 sentinel, so patterns never match across it.
SENTINEL = {q('w:' + t) for t in ('tab', 'br', 'cr', 'ptab', 'drawing', 'pict', 'object', 'sym', 'fldChar',
                                  'noBreakHyphen', 'softHyphen', 'footnoteReference', 'endnoteReference',
                                  'fldSimple')} | {q('m:oMath'), q('m:oMathPara'), q('mc:AlternateContent')}
SKIP = {q('w:' + t) for t in ('pPr', 'rPr', 'del', 'moveFrom', 'txbxContent', 'instrText', 'delText',
                              'delInstrText')} | {P}
LAYOUT_ONLY = {q('w:tab'), q('w:cr')}


def tokens(el):
    for c in el:
        if not isinstance(c.tag, str):
            continue
        if c.tag == T:
            yield c
        elif c.tag in SENTINEL:
            yield c
        elif c.tag not in SKIP:
            yield from tokens(c)


class Text:
    def __init__(self, p):
        self.segs, parts, pos = [], [], 0
        for el in tokens(p):
            s = (el.text or '') if el.tag == T else '\x00'
            self.segs.append((el, pos, pos + len(s)))
            parts.append(s)
            pos += len(s)
        self.text = ''.join(parts)

    def replace(self, start, end, repl):
        """Replace text[start:end] (start < end) in place. Callers apply edits right-to-left."""
        hit = [(el, s, e) for el, s, e in self.segs if s < end and e > start]
        if any(el.tag != T for el, _, _ in hit):
            return False
        first = True
        for el, s, _ in hit:
            txt = el.text or ''
            a, b = max(start, s) - s, min(end, s + len(txt)) - s
            el.text = txt[:a] + (repl if first else '') + txt[b:]
            first = False
            if el.text != el.text.strip():
                el.set(XML_SPACE, 'preserve')
        return True


def is_empty(p):
    """No text and no drawing/field/break content (layout-only tabs are allowed)."""
    for el in tokens(p):
        if el.tag == T:
            if (el.text or '').strip():
                return False
        elif el.tag == q('w:br'):
            if wv(el, 'type') in ('page', 'column'):
                return False
        elif el.tag not in LAYOUT_ONLY:
            return False
    return True


def is_bold(p, st):
    sid = p_style(p, st)
    seen = False
    for el in tokens(p):
        if el.tag != T or not (el.text or '').strip():
            continue
        seen = True
        r = el.getparent()
        v = is_on(r.find('w:rPr/w:b', NS))
        if v is None:
            rs = wv(r.find('w:rPr/w:rStyle', NS))
            v = st.bold(rs) if rs else None
        if v is None:
            v = st.bold(sid)
        if not (v if v is not None else st.doc_bold):
            return False
    return seen


SENTENCE_BREAK = re.compile(r'[.!?…]+["»”)]?\s+(?=[«"“(]?[A-ZА-ЯЁ0-9])')


def final_sentence(text):
    parts = SENTENCE_BREAK.split(text.strip())
    return parts[-1]


def ends_with_period(text):
    s = text.rstrip()
    return s.endswith('.') and not s.endswith('..')


def swap_final_period(p):
    t = Text(p)
    j = len(t.text.rstrip()) - 1
    return t.text[j] == '.' and t.replace(j, j + 1, ':')


def tie(p):
    pp = ppr(p)
    return set_flag(pp, 'keepNext', PPR_ORDER) + set_flag(pp, 'keepLines', PPR_ORDER)


def snippet(text, n=90):
    text = text.replace('\x00', ' ').strip()
    return text if len(text) <= n else text[:n - 1] + '…'


# ---------------------------------------------------------------------------------------------
# §8 Document sanitization (fixed sub-order: revisions -> comments -> metadata/whitespace).

CHANGE_RECORDS = [q('w:' + t) for t in ('pPrChange', 'rPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange',
                                        'sectPrChange', 'tblGridChange', 'tblPrExChange', 'numberingChange',
                                        'cellIns', 'cellDel', 'cellMerge')]
RANGE_MARKERS = [q('w:' + t) for t in ('moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart',
                                       'moveToRangeEnd', 'customXmlInsRangeStart', 'customXmlInsRangeEnd',
                                       'customXmlDelRangeStart', 'customXmlDelRangeEnd',
                                       'customXmlMoveFromRangeStart', 'customXmlMoveFromRangeEnd',
                                       'customXmlMoveToRangeStart', 'customXmlMoveToRangeEnd')]


def rule_accept_revisions(d):
    n = 0
    for root in d.word_parts:
        for el in list(root.iter(q('w:ins'), q('w:moveTo'))):
            parent = el.getparent()
            if parent.tag in (RPR, q('w:trPr')):
                remove(el)
            else:
                unwrap(el)
            n += 1
        for el in list(root.iter(q('w:del'), q('w:moveFrom'))):
            if el.getparent() is None:
                continue                    # already gone with an ancestor
            parent = el.getparent()
            n += 1
            if parent.tag == RPR and parent.getparent() is not None and parent.getparent().tag == PPR:
                # Deleted paragraph mark: the paragraph joins the next one.
                p = parent.getparent().getparent()
                remove(el)
                nxt = p.getnext()
                if nxt is not None and nxt.tag == P:
                    at = 1 if nxt.find(PPR) is not None else 0
                    for i, c in enumerate([c for c in p if c.tag != PPR]):
                        nxt.insert(at + i, c)
                    remove(p)
            elif parent.tag == RPR:
                remove(el)
            elif parent.tag == q('w:trPr'):
                remove(parent.getparent())  # deleted row
            else:
                remove(el)
        for el in list(root.iter(*CHANGE_RECORDS, *RANGE_MARKERS)):
            remove(el)
            n += 1
    settings = d.settings
    if settings is not None:
        for el in settings.findall(q('w:trackRevisions')):
            remove(el)
            n += 1
    return n


COMMENT_PARTS = ('comments.xml', 'commentsExtended.xml', 'commentsIds.xml', 'commentsExtensible.xml',
                 'people.xml')


def rule_strip_comments(d):
    n = 0
    for root in d.stories:
        for el in list(root.iter(q('w:commentRangeStart'), q('w:commentRangeEnd'), q('w:commentReference'))):
            parent = el.getparent()
            if el.tag == q('w:commentReference') and parent.tag == R and \
                    all(c is el or c.tag == RPR for c in parent):
                remove(parent)
            else:
                remove(el)
            n += 1
    folder = posixpath.dirname(d.main)
    for name in COMMENT_PARTS:
        n += d.pkg.delete(posixpath.join(folder, name))
    return n


def rule_metadata(d):
    n = 0
    core = d.xml((d.pkg.related('', 'core-properties') or [None])[0])
    if core is not None:
        for tag in ('dc:creator', 'cp:lastModifiedBy'):
            for el in core.findall(q(tag)):
                if el.text:
                    el.text = None
                    n += 1
        for el in core.findall(q('cp:revision')):
            remove(el)
            n += 1
    app = d.xml((d.pkg.related('', 'extended-properties') or [None])[0])
    if app is not None:
        for tag in ('Company', 'Manager', 'Template'):
            for el in app.findall(q('ep:' + tag)):
                if el.text:
                    el.text = None
                    n += 1
        for el in app.findall(q('ep:TotalTime')):
            if el.text != '0':
                el.text = '0'
                n += 1
    return n


def rule_macros(d):
    n = 0
    folder = posixpath.dirname(d.main)
    embed_ids = {rel.get('Id') for rel, t, _ in d.pkg.rels(d.main)
                 if t and t.startswith(folder + '/embeddings/')}
    for root in d.stories:
        for obj in list(root.iter(q('w:object'))):
            ids = {el.get(q('r:id')) for el in obj.iter(q('o:OLEObject'))}
            if ids & embed_ids:
                remove(obj)
                n += 1
    for name in list(d.pkg.names):
        if name in (folder + '/vbaProject.bin', folder + '/vbaData.xml') or name.startswith(folder + '/embeddings/'):
            n += d.pkg.delete(name)
    ct = d.xml(CONTENT_TYPES)
    for o in ct.findall(q('ct:Override')):
        if o.get('PartName') == '/' + d.main and 'macroEnabled' in o.get('ContentType', ''):
            o.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
            n += 1
    for dflt in ct.findall(q('ct:Default')):
        if dflt.get('Extension') == 'bin' and 'vbaProject' in dflt.get('ContentType', '') and \
                not any(x.endswith('.bin') for x in d.pkg.names):
            ct.remove(dflt)
            n += 1
    return n


def rule_hidden_text(d):
    n = 0
    for root in d.stories:
        for r in list(root.iter(R)):
            rpr = r.find(RPR)
            if rpr is None or not (is_on(rpr.find(q('w:vanish'))) or is_on(rpr.find(q('w:specVanish')))):
                continue
            if r.find(q('w:fldChar')) is not None or r.find(q('w:instrText')) is not None:
                d.findings['Hidden run kept (part of a field)'].append(snippet(Text(r).text or '(field)'))
                continue
            remove(r)
            n += 1
    return n


def rule_proofing(d):
    n = 0
    for root in d.stories:
        for el in list(root.iter(q('w:proofErr'), q('w:noProof'), q('w:lang'))):
            if el.tag == q('w:lang') and el.getparent().tag != RPR:
                continue
            remove(el)
            n += 1
    return n


def rule_protection(d):
    n = 0
    settings = d.settings
    if settings is not None:
        for el in settings.findall(q('w:documentProtection')) + settings.findall(q('w:writeProtection')):
            remove(el)
            n += 1
    return n


def rule_temp_bookmarks(d):
    n = 0
    for root in d.stories:
        ids = set()
        for el in list(root.iter(q('w:bookmarkStart'))):
            name = wv(el, 'name') or ''
            if name == '_GoBack' or name.startswith('OLE_LINK'):
                ids.add(wv(el, 'id'))
                remove(el)
                n += 1
        for el in list(root.iter(q('w:bookmarkEnd'))):
            if wv(el, 'id') in ids:
                remove(el)
    return n


STYLE_REF_TAGS = [q('w:' + t) for t in ('pStyle', 'rStyle', 'tblStyle', 'numStyleLink', 'styleLink',
                                        'clickAndTypeStyle', 'defaultTableStyle')]
FIELD_STYLE = re.compile(r'STYLEREF\s+"([^"]+)"|STYLEREF\s+(\S+)|\\t\s+"([^"]+)"')


def style_refs(d, st):
    """Style ids referenced from every part except styles.xml (incl. names in TOC \\t / STYLEREF)."""
    styles_part = d.xml(d.part('styles'))
    refs = set()
    for root in d.word_parts:
        if root is styles_part:
            continue
        for el in root.iter(*STYLE_REF_TAGS):
            refs.add(wv(el))
        for el in root.iter(q('w:instrText')):
            for m in FIELD_STYLE.finditer(el.text or ''):
                names = m.group(1) or m.group(2) or ''
                if m.group(3):
                    names = m.group(3)
                for name in names.split(',')[::2] if m.group(3) else [names]:
                    sid = st.by_name.get(name.strip()) or (name.strip() if name.strip() in st.by_id else None)
                    if sid:
                        refs.add(sid)
    refs.discard(None)
    return refs


def clean_styles(d):
    root = d.xml(d.part('styles'))
    if root is None:
        return 0
    st = Styles(root)
    keep = style_refs(d, st) | {sid for sid, s in st.by_id.items() if s.get(q('w:default')) in ('1', 'true')}
    todo = list(keep)
    while todo:
        s = st.by_id.get(todo.pop())
        if s is None:
            continue
        for tag in ('basedOn', 'link', 'next'):
            ref = wv(s.find(q('w:' + tag)))
            if ref and ref not in keep:
                keep.add(ref)
                todo.append(ref)
    backup = copy.deepcopy(root)
    removed = [s for sid, s in st.by_id.items() if sid not in keep]
    for s in removed:
        remove(s)
    # Verify: every reference still resolves, else revert this step.
    st2 = Styles(root)
    inner = {wv(el) for s in st2.by_id.values() for tag in ('basedOn', 'link', 'next')
             for el in s.findall(q('w:' + tag))}
    missing = (style_refs(d, st2) | inner) - set(st2.by_id) - {None}
    if missing - (style_refs(d, Styles(backup)) - set(Styles(backup).by_id)):
        d.pkg.trees[d.part('styles')] = backup
        d.findings['Orphan styles: reverted, references would break'].append(', '.join(sorted(missing)))
        return 0
    return len(removed)


def numbering_refs(d):
    numbering = d.xml(d.part('numbering'))
    refs = set()
    for root in d.word_parts:
        if root is numbering:
            continue
        for el in root.iter(q('w:numId')):
            if wv(el) not in (None, '0'):
                refs.add(wv(el))
    return refs


def clean_numbering(d):
    root = d.xml(d.part('numbering'))
    if root is None:
        return 0
    nums = {wv(n, 'numId'): n for n in root.findall(q('w:num'))}
    absts = {wv(a, 'abstractNumId'): a for a in root.findall(q('w:abstractNum'))}
    by_style_link = {wv(a.find(q('w:styleLink'))): aid for aid, a in absts.items()
                     if a.find(q('w:styleLink')) is not None}
    refs = numbering_refs(d)
    keep_nums = refs & set(nums)
    keep_abs = {wv(nums[n].find(q('w:abstractNumId'))) for n in keep_nums}
    todo = list(keep_abs)
    while todo:
        a = absts.get(todo.pop())
        link = wv(a.find(q('w:numStyleLink'))) if a is not None else None
        target = by_style_link.get(link)
        if target and target not in keep_abs:
            keep_abs.add(target)
            todo.append(target)
    backup = copy.deepcopy(root)
    n = 0
    for nid, el in nums.items():
        if nid not in keep_nums:
            remove(el)
            n += 1
    for aid, el in absts.items():
        if aid not in keep_abs:
            remove(el)
            n += 1
    # Verify: every referenced numId resolves num -> abstractNum, and every numStyleLink resolves.
    nums2 = {wv(x, 'numId'): x for x in root.findall(q('w:num'))}
    abs2 = {wv(a, 'abstractNumId'): a for a in root.findall(q('w:abstractNum'))}
    links2 = {wv(a.find(q('w:styleLink'))) for a in abs2.values()}
    broken = [r for r in refs if r in nums and (r not in nums2 or wv(nums2[r].find(q('w:abstractNumId'))) not in abs2)]
    broken += [wv(a.find(q('w:numStyleLink'))) for a in abs2.values()
               if a.find(q('w:numStyleLink')) is not None and wv(a.find(q('w:numStyleLink'))) not in links2
               and wv(a.find(q('w:numStyleLink'))) in by_style_link]
    if broken:
        d.pkg.trees[d.part('numbering')] = backup
        d.findings['Orphan numbering: reverted, references would break'].append(', '.join(map(str, broken)))
        return 0
    return n


def clean_rels_and_media(d):
    n = 0
    for rels_name in [x for x in d.pkg.names if x.endswith('.rels')]:
        for rel in list(d.pkg.xml(rels_name)):
            target = d.pkg.rel_target(rels_name, rel)
            if target is not None and target not in d.pkg.raw:
                remove(rel)
                n += 1
    targets = {d.pkg.rel_target(rn, rel) for rn in d.pkg.names if rn.endswith('.rels') for rel in d.pkg.xml(rn)}
    media = posixpath.dirname(d.main) + '/media/'
    for name in list(d.pkg.names):
        if name.startswith(media) and name not in targets:
            n += d.pkg.delete(name)
    return n


def rule_orphans(d):
    return clean_styles(d) + clean_numbering(d) + clean_rels_and_media(d)


def rule_strip_rsid(d):
    if not d.opts.strip_rsid:
        return 0
    n = 0
    settings = d.settings
    for el in settings.findall(q('w:rsids')):
        remove(el)
        n += 1
    for root in d.word_parts:
        for el in root.iter():
            for k in [k for k in el.attrib if k.startswith('{%s}rsid' % NS['w'])]:
                del el.attrib[k]
                n += 1
    n += ensure(settings, 'removePersonalInformation', SETTINGS_ORDER)[1]
    return n


# ---------------------------------------------------------------------------------------------
# §1-§4 Layout rules.

PGSZ = {'w': '11906', 'h': '16838', 'code': '9'}
PGMAR = {'top': '1134', 'right': '1134', 'bottom': '1134', 'left': '1134', 'header': '709', 'footer': '709',
         'gutter': '0'}


def rule_page_setup(d):
    n = 0
    for sp in d.xml(d.main).iter(q('w:sectPr')):
        szs, mars = sp.findall(q('w:pgSz')), sp.findall(q('w:pgMar'))
        if len(szs) == 1 and len(mars) == 1 and attrs(szs[0]) == PGSZ and attrs(mars[0]) == PGMAR:
            continue
        for el in szs + mars:
            sp.remove(el)
        insert_ordered(sp, mk('w:pgSz', PGSZ), SECTPR_ORDER)
        insert_ordered(sp, mk('w:pgMar', PGMAR), SECTPR_ORDER)
        n += 1
    return n


def tbl_pr(tbl):
    el = tbl.find(q('w:tblPr'))
    if el is None:
        el = mk('w:tblPr')
        tbl.insert(0, el)
    return el


def rule_table_width(d):
    n = 0
    for tbl in d.xml(d.main).iter(TBL):
        pr = tbl_pr(tbl)
        changed = set_attrs(ensure(pr, 'tblW', TBLPR_ORDER)[0], {'w': '5000', 'type': 'pct'})
        changed += set_attrs(ensure(pr, 'tblLayout', TBLPR_ORDER)[0], {'type': 'autofit'})
        ind = pr.find(q('w:tblInd'))
        if ind is not None and wv(ind, 'w') not in (None, '0'):
            changed += set_attrs(ind, {'w': '0', 'type': 'dxa'})
        n += bool(changed)
    d.stats['tables'] = sum(1 for _ in d.xml(d.main).iter(TBL))
    return n


def tr_pr(tr):
    el = tr.find(q('w:trPr'))
    if el is None:
        el = mk('w:trPr')
        tr.insert(1 if tr.find(q('w:tblPrEx')) is not None else 0, el)
    return el


def rule_table_breaks(d):
    n = 0
    for i, tbl in enumerate(d.xml(d.main).iter(TBL), 1):
        rows = tbl.findall(q('w:tr'))
        if not rows:
            continue
        pr = tbl_pr(tbl)
        look, created = ensure(pr, 'tblLook', TBLPR_ORDER)
        changed = created
        if wv(look, 'firstRow') not in ('1', 'true'):
            look.set(q('w:firstRow'), '1')
            changed = 1
        if wv(look) is not None and not int(wv(look), 16) & 0x20:
            look.set(q('w:val'), '%04X' % (int(wv(look), 16) | 0x20))
            changed = 1
        header = tr_pr(rows[0])
        if header.find(q('w:tblHeader')) is None or not is_on(header.find(q('w:tblHeader'))):
            first_cell = ' | '.join(Text(p).text.strip() for p in rows[0].iter(P))
            d.log['Table first row newly set as repeating header (check it is a header)'].append(
                f'table {i}: {snippet(first_cell, 70)}')
        changed += set_flag(header, 'tblHeader', TRPR_ORDER)
        for tr in rows:
            changed += set_flag(tr_pr(tr), 'cantSplit', TRPR_ORDER)
        n += bool(changed)
    return n


def rule_list_pagination(d):
    n = items = runs = leadins = 0
    for root in d.stories:
        for cont in containers(root):
            bl = blocks(cont)
            for i, b in enumerate(bl):
                if b.tag != P:
                    continue
                nxt = bl[i + 1] if i + 1 < len(bl) else None
                nxt_list = nxt is not None and nxt.tag == P and is_list(nxt)
                if is_list(b):
                    items += 1
                    runs += not nxt_list
                    n += set_flag(ppr(b), 'keepLines', PPR_ORDER)
                    if nxt_list:
                        n += set_flag(ppr(b), 'keepNext', PPR_ORDER)
                elif nxt_list and Text(b).text.strip() and b.find('w:pPr/w:sectPr', NS) is None:
                    leadins += 1
                    n += tie(b)
    d.stats.update({'list items': items, 'lists': runs, 'list lead-ins': leadins})
    return n


# ---------------------------------------------------------------------------------------------
# §5 / §10 Lead-in punctuation.

COUNT_WORDS = [(k, re.compile(r'\b(?:%s)\b' % v, re.I)) for k, v in [
    (2, r'дв(?:а|е|ух|ум|умя|ое)|two'),
    (3, r'тр(?:и|ёх|ех|ём|ем|емя|ое)|three'),
    (4, r'четыр(?:е|ёх|ех|ём|ем|ьмя)|четверо|four'),
    (5, r'пят(?:ь|и|ью|еро)|five'),
    (6, r'шест(?:ь|и|ью|еро)|six'),
    (7, r'сем(?:ь|и|ью|еро)|seven'),
    (8, r'вос(?:емь|ьми|емью|ьмеро)|eight'),
    (9, r'девят(?:ь|и|ью)|nine'),
    (10, r'десят(?:ь|и|ью)|ten'),
    (11, r'одиннадцат(?:ь|и|ью)|eleven'),
    (12, r'двенадцат(?:ь|и|ью)|twelve'),
]]
DIGIT_COUNT = re.compile(r'\b(\d{1,2})\b')
ENUMERATIVE = re.compile(
    r'состо(?:ит|ят) из|включа(?:ет|ют)|охватыва(?:ет|ют)|дел(?:ит|ят)ся на|разделя(?:ет|ют)ся на|'
    r'различа(?:ем|ют)|выделя(?:ем|ют)|следующ|а именно|относятся|сводятся к|перечисл|'
    r'\bconsists? of|\bcompris(?:es|e)\b|\binclud(?:es|e)\b|\bcovers?\b|(?:is|are) divided into|'
    r'\bwe distinguish|\bthe following\b|\bas follows\b|\bnamely\b|\bfalls? into\b', re.I)
SUMMARY_OPENER = re.compile(
    r'^(?:таким образом|итак|в итоге|в результате|следовательно|thus|therefore|in summary|overall|'
    r'in short|as a result)\b', re.I)
ABBREVIATION_END = re.compile(r'(?:\bт\.\s?[дпе]|\bдр|\bсм|\betc|\be\.g|\bi\.e|\bvs)\.$', re.I)


def leadin_id(text):
    return hashlib.sha1(text.strip().encode()).hexdigest()[:8]


def classify_leadin(text, n_items):
    """(convert?, reason) for a '.'-ending paragraph right before a list of n_items (§10a)."""
    sentence = final_sentence(text)
    if ':' in sentence:
        return False, 'colon guard: final sentence already has a colon'
    if ABBREVIATION_END.search(text.rstrip()):
        return False, 'review: ends with an abbreviation'
    if SUMMARY_OPENER.search(sentence):
        return False, 'review: reads as a summary of the previous block'
    # A count may sit in an earlier sentence ("Three anti-patterns recur. Each has its own fix.")
    counts = {k for k, rx in COUNT_WORDS if rx.search(text)}
    counts |= {int(m) for m in DIGIT_COUNT.findall(text)}
    if n_items in counts:
        return True, f'count agreement ({n_items} items)'
    if ENUMERATIVE.search(sentence):
        return True, 'enumerative phrasing'
    return False, 'review: no deterministic signal'


def rule_leadin_dot(d):
    n = 0
    st = d.styles()
    for cont in containers(d.xml(d.main)):
        bl = blocks(cont)
        for i, b in enumerate(bl[:-1]):
            nxt = bl[i + 1]
            if b.tag != P or is_list(b) or nxt.tag != P or not is_list(nxt) or d.heading_level(b, st):
                continue
            text = Text(b).text
            if not text.strip() or not ends_with_period(text):
                continue
            n_items = 0
            for x in bl[i + 1:]:
                if x.tag != P or not is_list(x):
                    break
                n_items += 1
            lid = leadin_id(text)
            convert, reason = classify_leadin(text, n_items)
            if lid in d.opts.leadins:
                convert, reason = True, 'accepted via --leadin'
            if convert and swap_final_period(b):
                n += 1 + tie(b)
                d.log['Lead-in before a list: "." -> ":"'].append(f'{snippet(text)} — {reason}')
            elif reason.startswith('review'):
                d.findings['Lead-in candidates left for review (accept with --leadin ID)'].append(
                    f'`{lid}` {snippet(text, 140)} → list of {n_items} — {reason}')
            else:
                d.findings['Lead-ins kept with "."'].append(f'{snippet(text)} — {reason}')
    return n


def rule_bold_leadin(d):
    """§5b: single bold sentence right after a heading 1-3, ending in '.', followed by text or a table."""
    n = 0
    st = d.styles()
    for cont in containers(d.xml(d.main)):
        bl = blocks(cont)
        for i in range(1, len(bl) - 1):
            prev, cur, nxt = bl[i - 1], bl[i], bl[i + 1]
            if prev.tag != P or cur.tag != P or nxt.tag not in (P, TBL) or is_list(cur):
                continue
            lvl = d.heading_level(prev, st)
            if not lvl or lvl > 3 or d.heading_level(cur, st):
                continue
            text = Text(cur).text
            if not ends_with_period(text) or SENTENCE_BREAK.search(text.strip()) \
                    or (nxt.tag == P and not Text(nxt).text.strip()) or not is_bold(cur, st):
                continue
            if ':' in final_sentence(text):
                changed = tie(cur)
                if changed:
                    d.log['Bold lead-in tied, "." kept (colon guard)'].append(snippet(text))
            else:
                changed = swap_final_period(cur) + tie(cur)
                d.log['Bold lead-in after heading: "." -> ":"'].append(snippet(text))
            n += changed
    return n


# ---------------------------------------------------------------------------------------------
# §6 Text sanitization, §7 structure.

RANGE_SEP = re.compile(r'(?<=\S)(?: +- *| *- +)(?=\S)')
NUM_TOKEN = re.compile(r'^[$€₽£¥]?[A-Za-zА-Яа-яЁё]{0,3}\d[\d.,]*(?:%|[$€₽£¥]|[A-Za-zА-Яа-яЁё]{1,6}\.?)?$')


def range_dashes(text):
    for m in RANGE_SEP.finditer(text):
        left = re.split(r'[\s\x00]', text[:m.start()])[-1].lstrip('(«"“')
        right = re.split(r'[\s\x00]', text[m.end():])[0].rstrip(',.;:)»"”!?')
        if NUM_TOKEN.match(left) and NUM_TOKEN.match(right) and not left.startswith(('-', '−')):
            yield m.start(), m.end(), '–'


def apply(p, finder):
    t = Text(p)
    edits = list(finder(t.text))
    for start, end, repl in reversed(edits):
        t.replace(start, end, repl)
    return len(edits)


def trims(text):
    m = re.search(r'[ \t]+$', text)
    if m:
        yield m.start(), m.end(), ''
    m = re.match(r'^[ \t]+', text)
    if m and m.end() < len(text):
        yield m.start(), m.end(), ''


def double_spaces(text):
    for m in re.finditer(r' {2,}', text):
        yield m.start(), m.end(), ' '


def rule_text(d):
    counts = defaultdict(int)
    for root in d.stories:
        for p in root.iter(P):
            counts['range dashes'] += apply(p, range_dashes)
            counts['trims'] += apply(p, lambda s: sorted(trims(s)))
            counts['double spaces'] += apply(p, double_spaces)
    d.stats.update({'text: ' + k: v for k, v in counts.items()})
    return sum(counts.values())


def protected(p):
    return p.find('w:pPr/w:sectPr', NS) is not None or p.find('w:pPr/w:pageBreakBefore', NS) is not None \
        or any(True for _ in p.iter(q('w:bookmarkStart'), q('w:bookmarkEnd')))


def rule_blank_lines(d):
    n = 0
    for root in d.stories:
        for cont in containers(root):
            run = []
            for b in blocks(cont) + [None]:
                if b is not None and b.tag == P and is_empty(b):
                    run.append(b)
                    continue
                if len(run) >= 2 and len([p for p in run if protected(p)] or run[:1]) < len(run):
                    keep = [p for p in run if protected(p)] or run[:1]
                    for p in run:
                        if p not in keep:
                            remove(p)
                            n += 1
                    after = snippet(Text(b).text, 60) if b is not None and b.tag == P else \
                        'table' if b is not None else 'end of ' + local(cont)
                    d.log['Runs of empty paragraphs collapsed'].append(f'{len(run)} → {len(keep)}, before: {after}')
                run = []
        # empty lines made with line breaks: repeated, leading or trailing breaks inside a paragraph
        for p in root.iter(P):
            toks = [el for el in tokens(p) if not (el.tag == T and not (el.text or '').strip())]
            for i, el in enumerate(toks):
                if el.tag != q('w:br') or wv(el, 'type') not in (None, 'textWrapping'):
                    continue
                prev_br = i > 0 and toks[i - 1].tag == q('w:br') and wv(toks[i - 1], 'type') in (None, 'textWrapping')
                if i == 0 or i == len(toks) - 1 or prev_br:
                    remove(el)
                    d.log['Empty lines made with line breaks removed'].append(snippet(Text(p).text, 60))
                    n += 1
    return n


def rule_empty_headings(d):
    """§7a: report only — an empty heading may be a deliberate anchor."""
    st = d.styles()
    for p in d.xml(d.main).iter(P):
        lvl = d.heading_level(p, st)
        if lvl and lvl <= 3 and not Text(p).text.replace('\x00', '').strip():
            marks = [wv(b, 'name') for b in p.iter(q('w:bookmarkStart'))]
            d.findings['Empty headings (not removed)'].append(
                f'H{lvl} style {p_style(p, st)}' + (f', bookmarks: {", ".join(marks)}' if marks else ''))
    return 0


# ---------------------------------------------------------------------------------------------
# §11 Styles: every paragraph's look comes from its style, not from per-editor direct formatting.
# One-way sub-order (all within one pass): assign styles (11.1–11.3) -> promote the dominant direct
# formatting into each style (11.4) -> strip direct formatting (11.5) -> house style (11.6, 11.7).

RPR_ORDER = ('rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint '
             'noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u '
             'effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath').split()
TCPR_ORDER = ('cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText '
              'vAlign hideMark headers cellIns cellDel cellMerge tcPrChange').split()
STYLE_ORDER = ('name aliases basedOn next link autoRedefine hidden uiPriority semiHidden unhideWhenUsed '
               'qFormat locked personal personalCompose personalReply rsid pPr rPr tblPr trPr tcPr '
               'tblStylePr').split()
ORDERS.update({'rPr': RPR_ORDER, 'tcPr': TCPR_ORDER, 'style': STYLE_ORDER})

# House style (§11.6, §11.7).
BODY_JC = 'both'
BODY_SPACING = {'before': '120', 'after': '120'}       # 6 pt before and after
CELL_MARGIN = '144'                                      # 0.1" on every side
KEEP_FONTS = re.compile(r'^(Symbol|Wingdings.*|Webdings|Cambria Math|Consolas|Courier New|Menlo|'
                        r'Source Code Pro|JetBrains Mono)$')
# Direct run properties that only carry an editor's look: stripped wherever they appear.
COSMETIC_R = {'rFonts', 'sz', 'szCs', 'color', 'highlight', 'kern', 'w', 'spacing', 'noProof', 'lang'}
# Direct run properties with meaning (emphasis, sub/superscript): kept unless the style already says so.
SEMANTIC_R = {'b', 'bCs', 'i', 'iCs', 'u', 'strike', 'dstrike', 'caps', 'smallCaps', 'vertAlign'}
# Direct paragraph properties that only carry an editor's look.
COSMETIC_P = {'spacing', 'ind', 'jc', 'contextualSpacing', 'suppressAutoHyphens'}
# Pagination flags: meaningful, but redundant where the style already says the same.
FLOW_P = {'keepNext', 'keepLines', 'pageBreakBefore', 'widowControl'}


def rvalue(el):
    """Comparable value of an rPr/pPr child."""
    name = local(el)
    if name == 'rFonts':
        return el.get(q('w:ascii')) or ('theme:' + el.get(q('w:asciiTheme'))
                                        if el.get(q('w:asciiTheme')) else None)
    if name in SEMANTIC_R and name not in ('u', 'vertAlign'):
        return is_on(el)
    if name == 'color':
        return (el.get(q('w:val')) or '').upper()
    return tuple(sorted(attrs(el).items()))


def style_prop(st, sid, where, name):
    """The element `name` under w:pPr / w:rPr resolved along the basedOn chain (None if unset)."""
    for s in st.chain(sid):
        el = s.find(f'w:{where}/w:{name}', NS)
        if el is not None:
            return el
    return None


def style_el(st, sid, where):
    s = st.by_id[sid]
    el = s.find(q('w:' + where))
    if el is None:
        el = insert_ordered(s, mk('w:' + where), STYLE_ORDER)
    return el


def font_name(rfonts):
    return rfonts.get(q('w:ascii')) or rfonts.get(q('w:hAnsi')) or ''


def keep_font(el):
    return bool(KEEP_FONTS.match(font_name(el)))


def text_runs(p):
    """(run, text) for runs directly carrying text in this paragraph (hyperlinks included)."""
    out = []
    for r in p.iter(R):
        if r.getparent().tag in (q('w:del'), q('w:moveFrom')):
            continue
        txt = ''.join(t.text or '' for t in r.findall(T))
        if txt.strip():
            out.append((r, txt))
    return out


def body_paragraphs(d):
    """Paragraphs directly in the body (not tables, text boxes, headers, notes)."""
    body = d.xml(d.main).find(q('w:body'))
    return [p for p in body.iter(P) if p.getparent() is body or
            (p.getparent().tag == q('w:sdtContent') and p.getparent().getparent().getparent() is body)]


def in_table(p):
    return any(a.tag == q('w:tc') for a in p.iterancestors())


def roles(d, st):
    """Which styles play body text, table cell and table header roles (by amount of text)."""
    body, cell, head = Counter(), Counter(), Counter()
    for p in body_paragraphs(d):
        if d.heading_level(p, st) is None and not is_list(p):
            body[p_style(p, st)] += len(Text(p).text.strip())
    for tbl in d.xml(d.main).iter(TBL):
        for i, tr in enumerate(tbl.findall(TR)):
            for p in tr.iter(P):
                (head if i == 0 else cell)[p_style(p, st)] += len(Text(p).text.strip())
    top = lambda c: c.most_common(1)[0][0] if c else None
    return top(body), top(cell), top(head)


def rule_unstyled(d):
    """§11.1 A body paragraph left in the default style gets the body text style."""
    st = d.styles()
    if 'text language' not in d.stats:
        d.stats['text language'] = text_language(d)
    body, _, _ = roles(d, st)
    n = 0
    if body is None or body == st.default_p:
        return 0
    for p in body_paragraphs(d):
        if wv(p.find('w:pPr/w:pStyle', NS)) is None and Text(p).text.replace('\x00', '').strip() \
                and d.heading_level(p, st) is None and p.find('w:pPr/w:framePr', NS) is None:
            pp = ppr(p)
            pp.insert(0, mk('w:pStyle', {'val': body}))
            d.log['Unstyled paragraphs given the body style'].append(snippet(Text(p).text))
            n += 1
    return n


NUMBERED = re.compile(r'^\s*(\d+(?:\.\d+)*)\.?\s+\S')


def rule_heading_gaps(d):
    """§11.2 A paragraph whose number fills a gap between two same-level headings becomes one."""
    st = d.styles()
    ps = body_paragraphs(d)
    heads = []
    for i, p in enumerate(ps):
        m = NUMBERED.match(Text(p).text)
        lvl = d.heading_level(p, st)
        if m and lvl:
            heads.append((i, tuple(int(x) for x in m.group(1).split('.')), p_style(p, st)))
    n = 0
    for (i, a, sa), (j, b, sb) in zip(heads, heads[1:]):
        if sa != sb or len(a) != len(b) or a[:-1] != b[:-1] or b[-1] != a[-1] + 2:
            continue
        want = a[:-1] + (a[-1] + 1,)
        for p in ps[i + 1:j]:
            text = Text(p).text.replace('\x00', '').strip()
            m = NUMBERED.match(text)
            if m and tuple(int(x) for x in m.group(1).split('.')) == want and len(text) < 160 \
                    and not text.endswith(('.', ':', ';')) and d.heading_level(p, st) is None:
                pp = ppr(p)
                old = pp.find(q('w:pStyle'))
                if old is not None:
                    remove(old)
                pp.insert(0, mk('w:pStyle', {'val': sa}))
                d.log['Numbered headings set in a body style, restored as headings'].append(snippet(text))
                n += 1
    return n


CAPTION_WORDS = re.compile(r'^\s*(Рисунок|Рис\.|Таблица|Figure|Fig\.|Table)\s*\d', re.I)


def is_italic(r, st, sid):
    v = is_on(r.find('w:rPr/w:i', NS))
    if v is None:
        v = is_on(style_prop(st, sid, 'rPr', 'i'))
    return bool(v)


def next_block(p):
    """The next table, picture or non-empty paragraph after p (empty paragraphs are skipped)."""
    nxt = p.getnext()
    while nxt is not None:
        if nxt.tag == TBL:
            return 'table'
        if nxt.tag == P:
            if nxt.find('.//w:drawing', NS) is not None or nxt.find('.//w:pict', NS) is not None:
                return 'picture'
            if Text(nxt).text.replace('\x00', '').strip():
                return 'text'
        nxt = nxt.getnext()
    return None


def rule_captions(d):
    """§11.3 Captions and block titles get their own styles.
    - a paragraph starting "Рисунок N" / "Figure N" is a caption: the caption style;
    - a short, fully italic line without final punctuation right above a table or picture is a block
      title: a title style derived from the caption style (left-aligned)."""
    st = d.styles()
    names = {sid: (wv(s.find(q('w:name'))) or '') for sid, s in st.by_id.items()}
    use = Counter(p_style(p, st) for p in body_paragraphs(d))
    caps = [sid for sid in use if 'caption' in names.get(sid, '').lower() and 'title' not in names.get(sid, '').lower()]
    if not caps:
        return 0
    cap = max(caps, key=lambda sid: use[sid])
    title = cap + 'Title'
    body, _, _ = roles(d, st)
    n = 0
    for p in body_paragraphs(d):
        sid = p_style(p, st)
        text = Text(p).text.replace('\x00', '').strip()
        if not text or d.heading_level(p, st) is not None or is_list(p) or sid not in (body, cap):
            continue
        if CAPTION_WORDS.match(text):
            want, label = cap, 'Captions set in another style, given the caption style'
        elif len(text) < 120 and not text.endswith(('.', ':', ';', '!', '?')) \
                and next_block(p) in ('table', 'picture') \
                and all(is_italic(r, st, sid) for r, _ in text_runs(p)):
            want, label = title, 'Block titles above tables and pictures, given the title style'
        else:
            continue
        if sid == want:
            continue
        if want == title and title not in st.by_id:
            s = mk('w:style', {'type': 'paragraph', 'customStyle': '1', 'styleId': title})
            s.append(mk('w:name', {'val': names[cap] + ' — Title'}))
            s.append(mk('w:basedOn', {'val': cap}))
            s.append(mk('w:qFormat'))
            pp = s.makeelement(q('w:pPr'))
            pp.append(mk('w:keepNext'))
            pp.append(mk('w:jc', {'val': 'left'}))
            s.append(pp)
            st.by_id[cap].addnext(s)
            st.by_id[title] = s
            d.log['Styles created'].append(f'{title}: block titles, based on {cap}')
        pp = ppr(p)
        old = pp.find(q('w:pStyle'))
        if old is not None:
            remove(old)
        pp.insert(0, mk('w:pStyle', {'val': want}))
        d.log[label].append(snippet(text))
        n += 1
    return n


def has_picture(p):
    return p.find('.//w:drawing', NS) is not None or p.find('.//w:pict', NS) is not None


def rule_figure_keep(d):
    """§11.9 A caption stays on the page of its picture or table: a picture paragraph followed by a
    caption keeps with next (so do empty paragraphs between them); a caption followed by a picture or
    table keeps with next."""
    st = d.styles()
    n = 0
    for p in body_paragraphs(d):
        text = Text(p).text.replace('\x00', '').strip()
        if has_picture(p) and not text:
            chain, nxt = [p], p.getnext()
            while nxt is not None and nxt.tag == P and not has_picture(nxt) \
                    and not Text(nxt).text.replace('\x00', '').strip():
                chain.append(nxt)
                nxt = nxt.getnext()
            if nxt is not None and nxt.tag == P and CAPTION_WORDS.match(Text(nxt).text.replace('\x00', '')):
                for x in chain:
                    if not is_on(style_prop(st, p_style(x, st), 'pPr', 'keepNext')):
                        n += set_flag(ppr(x), 'keepNext', PPR_ORDER)
        elif CAPTION_WORDS.match(text) and next_block(p) in ('table', 'picture') \
                and not is_on(style_prop(st, p_style(p, st), 'pPr', 'keepNext')):
            n += set_flag(ppr(p), 'keepNext', PPR_ORDER)
    return n


def rule_promote(d):
    """§11.4 The direct formatting most of a style's text carries becomes part of that style."""
    st = d.styles()
    by_style = defaultdict(list)
    for root in d.stories:
        for p in root.iter(P):
            by_style[p_style(p, st)].append(p)
    n = 0
    for sid, ps in by_style.items():
        if sid not in st.by_id:
            continue
        # run properties, weighted by characters; runs with a character style are its business
        weight, votes = 0, defaultdict(Counter)
        for p in ps:
            for r, txt in text_runs(p):
                if r.find('w:rPr/w:rStyle', NS) is not None:
                    continue
                weight += len(txt)
                rpr = r.find(RPR)
                for name in ('rFonts', 'sz', 'color', 'b', 'i'):
                    el = rpr.find(q('w:' + name)) if rpr is not None else None
                    votes[name][None if el is None else etree.tostring(el)] += len(txt)
        for name, c in votes.items():
            val, cnt = c.most_common(1)[0]
            if val is None or cnt * 2 <= weight:
                continue
            el = etree.fromstring(val)
            if name == 'rFonts' and keep_font(el):
                continue
            cur = style_prop(st, sid, 'rPr', name)
            if cur is not None and rvalue(cur) == rvalue(el):
                continue
            rpr = style_el(st, sid, 'rPr')
            old = rpr.find(q('w:' + name))
            if old is not None:
                remove(old)
            insert_ordered(rpr, el, RPR_ORDER)
            if name == 'sz':
                for x in rpr.findall(q('w:szCs')):
                    remove(x)
                insert_ordered(rpr, mk('w:szCs', attrs(el)), RPR_ORDER)
            d.log['Formatting promoted into styles'].append(f'{sid}: {name} {rvalue(el)}')
            n += 1
        # paragraph spacing, per attribute, over the style's non-empty paragraphs
        texts = [p for p in ps if Text(p).text.replace('\x00', '').strip()]
        sv = defaultdict(Counter)
        for p in texts:
            sp = p.find('w:pPr/w:spacing', NS)
            got = attrs(sp) if sp is not None else {}
            for k in ('before', 'after', 'line', 'lineRule'):
                sv[k][got.get(k)] += 1
        for k, c in sv.items():
            val, cnt = c.most_common(1)[0]
            if val is None or cnt * 2 <= len(texts):
                continue
            cur = style_prop(st, sid, 'pPr', 'spacing')
            if cur is not None and cur.get(q('w:' + k)) == val:
                continue
            sp, _ = ensure(style_el(st, sid, 'pPr'), 'spacing', PPR_ORDER)
            sp.set(q('w:' + k), val)
            d.log['Formatting promoted into styles'].append(f'{sid}: spacing {k}={val}')
            n += 1
        # alignment, over the style's non-empty paragraphs
        jc = Counter(wv(p.find('w:pPr/w:jc', NS)) for p in texts)
        if jc and st.outline(sid) is None:      # headings stay left-aligned
            val, cnt = jc.most_common(1)[0]
            cur = wv(style_prop(st, sid, 'pPr', 'jc'))
            if val is not None and cnt * 2 > len(texts) and cur != val:
                set_attrs(ensure(style_el(st, sid, 'pPr'), 'jc', PPR_ORDER)[0], {'val': val})
                d.log['Formatting promoted into styles'].append(f'{sid}: jc {val}')
                n += 1
    return n


def rule_strip_direct(d):
    """§11.5 Remove direct formatting: looks always, emphasis only where the style already has it."""
    st = d.styles()
    n = 0
    for root in d.stories:
        for p in root.iter(P):
            sid = p_style(p, st)
            pp = p.find(PPR)
            seen = set()
            for r, _ in text_runs(p):
                rpr = r.find(RPR)
                if rpr is not None:
                    for el in list(rpr):
                        if local(el) in ('rFonts', 'sz', 'color') and \
                                not (local(el) == 'rFonts' and keep_font(el)):
                            cur = style_prop(st, sid, 'rPr', local(el))
                            if cur is None or rvalue(cur) != rvalue(el):
                                seen.add(local(el))
            if seen and root is d.xml(d.main) and re.fullmatch(r'toc \d', wv(st.by_id[sid].find(q('w:name')))
                                                                if sid in st.by_id else '') :
                d.stats['TOC entries with mixed fonts'] = d.stats.get('TOC entries with mixed fonts', 0) + 1
            elif seen and root is d.xml(d.main):
                d.log['Formatting artifacts removed (font / size / colour differing from the style)'].append(
                    f'{sid}: {", ".join(sorted(seen))} — {snippet(Text(p).text, 70)}')
            for r in p.iter(R):
                rpr = r.find(RPR)
                if rpr is None:
                    continue
                rs = wv(rpr.find(q('w:rStyle')))
                for el in list(rpr):
                    name = local(el)
                    if name in COSMETIC_R:
                        if name == 'rFonts' and keep_font(el):
                            continue
                        if name == 'lang' and rs:
                            continue
                        remove(el)
                        n += 1
                    elif name in SEMANTIC_R and not rs:
                        cur = style_prop(st, sid, 'rPr', name)
                        if (cur is not None and rvalue(cur) == rvalue(el)) or \
                                (cur is None and rvalue(el) is False):
                            remove(el)
                            n += 1
                if len(rpr) == 0:
                    remove(rpr)
            if pp is not None:
                mark = pp.find(RPR)
                if mark is not None:
                    for el in list(mark):
                        if local(el) in COSMETIC_R | SEMANTIC_R:
                            remove(el)
                            n += 1
                for el in list(pp):
                    name = local(el)
                    if name in FLOW_P:
                        cur = style_prop(st, sid, 'pPr', name)
                        if is_on(el) == bool(is_on(cur)):
                            remove(el)
                            n += 1
                        continue
                    if name not in COSMETIC_P:
                        continue
                    if name == 'ind' and is_list(p):
                        kept = {k: v for k, v in attrs(el).items() if k not in ('right', 'end')}
                        if kept != attrs(el):
                            set_attrs(el, kept)
                            n += 1
                        continue
                    remove(el)
                    n += 1
        for el in list(root.iter(q('w:proofErr'))):
            remove(el)
            n += 1
    return n


def dark_grey(color):
    try:
        r, g, b = (int(color[i:i + 2], 16) for i in (0, 2, 4))
    except (TypeError, ValueError):
        return False
    return max(r, g, b) < 0x50 and max(r, g, b) - min(r, g, b) < 0x20


def rule_house_style(d):
    """§11.6 Body text: justified, 6 pt before/after, black, hyphenated; standard (theme) fonts."""
    st = d.styles()
    root = d.xml(d.part('styles'))
    body, _, _ = roles(d, st)
    n = 0
    if body in st.by_id:
        pp = style_el(st, body, 'pPr')
        n += set_attrs(ensure(pp, 'jc', PPR_ORDER)[0], {'val': BODY_JC})
        sp, _ = ensure(pp, 'spacing', PPR_ORDER)
        for k, v in BODY_SPACING.items():
            if sp.get(q('w:' + k)) != v:
                sp.set(q('w:' + k), v)
                n += 1
        for el in pp.findall(q('w:suppressAutoHyphens')):
            remove(el)
            n += 1
    # exact line height clips pictures and large glyphs: single spacing instead
    for sp in root.iter(q('w:spacing')):
        if sp.getparent().tag == PPR and sp.get(q('w:lineRule')) == 'exact':
            sp.set(q('w:line'), '240')
            sp.set(q('w:lineRule'), 'auto')
            n += 1
    heading_styles = {sid for sid in st.by_id if (st.outline(sid) is not None and st.outline(sid) < 9)}
    for s in root.iter(q('w:style')):
        sid = s.get(q('w:styleId'))
        for rf in list(s.iter(q('w:rFonts'))):
            if keep_font(rf):
                continue
            want = 'majorHAnsi' if sid in heading_styles else None
            if want is None:
                remove(rf)
                n += 1
            elif attrs(rf) != {'asciiTheme': want, 'hAnsiTheme': want, 'eastAsiaTheme': 'majorEastAsia',
                               'cstheme': 'majorBidi'}:
                set_attrs(rf, {'asciiTheme': want, 'hAnsiTheme': want, 'eastAsiaTheme': 'majorEastAsia',
                               'cstheme': 'majorBidi'})
                n += 1
        for c in list(s.iter(q('w:color'))):
            if c.getparent().tag == RPR and (wv(c) == '000000' or dark_grey(wv(c))) and not c.get(q('w:themeColor')):
                remove(c)
                n += 1
    # headings and block titles are black by default; a colour scheme comes only from --scheme
    scheme = scheme_styles(d, st)
    titles = heading_styles | {sid for sid in st.by_id if sid.endswith('Title')}
    titles |= {wv(st.by_id[sid].find(q('w:link'))) for sid in titles} - {None}
    for sid in titles - set(scheme):
        for c in list(st.by_id[sid].iter(q('w:color'))):
            if c.getparent().tag == RPR:
                remove(c)
                n += 1
    for story in d.stories:
        for p in story.iter(P):
            if p_style(p, st) in titles:
                for c in list(p.iter(q('w:color'))):
                    if c.getparent().tag == RPR:
                        remove(c)
                        n += 1
    # headings never end a page: keep with next; a level-1 heading starts a new page
    for sid in heading_styles:
        pp = style_el(st, sid, 'pPr')
        n += set_flag(pp, 'keepNext', PPR_ORDER)
        if st.outline(sid) == 0:
            n += set_flag(pp, 'pageBreakBefore', PPR_ORDER)
    # heading styles whose chain lost every font: give the top heading style the theme heading font
    for sid in heading_styles:
        if style_prop(st, sid, 'rPr', 'rFonts') is None:
            insert_ordered(style_el(st, sid, 'rPr'), mk('w:rFonts', {
                'asciiTheme': 'majorHAnsi', 'hAnsiTheme': 'majorHAnsi', 'eastAsiaTheme': 'majorEastAsia',
                'cstheme': 'majorBidi'}), RPR_ORDER)
            n += 1
    # TOC levels share one size: the most common among the levels that set one
    tocs = [sid for sid, s in st.by_id.items() if re.fullmatch(r'toc \d', wv(s.find(q('w:name'))) or '')]
    sizes = Counter(wv(style_prop(st, sid, 'rPr', 'sz')) for sid in tocs if style_prop(st, sid, 'rPr', 'sz') is not None)
    if sizes:
        size = sizes.most_common(1)[0][0]
        for sid in tocs:
            if wv(style_prop(st, sid, 'rPr', 'sz')) != size:
                rpr = style_el(st, sid, 'rPr')
                for tag in ('sz', 'szCs'):
                    for x in rpr.findall(q('w:' + tag)):
                        remove(x)
                    insert_ordered(rpr, mk('w:' + tag, {'val': size}), RPR_ORDER)
                n += 1
    # document defaults: theme body font, black, the language most of the text is in
    defaults = root.find('w:docDefaults/w:rPrDefault/w:rPr', NS)
    if defaults is not None:
        rf = defaults.find(q('w:rFonts'))
        want = {'asciiTheme': 'minorHAnsi', 'hAnsiTheme': 'minorHAnsi', 'eastAsiaTheme': 'minorEastAsia',
                'cstheme': 'minorBidi'}
        if rf is None:
            rf = insert_ordered(defaults, mk('w:rFonts'), RPR_ORDER)
        n += set_attrs(rf, want)
        for c in defaults.findall(q('w:color')):
            remove(c)
            n += 1
        lang = d.stats.get('text language')
        if lang:
            el, _ = ensure(defaults, 'lang', RPR_ORDER)
            if wv(el) != lang:
                el.set(q('w:val'), lang)
                n += 1
    settings = d.settings
    if settings is not None:
        n += set_flag(settings, 'autoHyphenation', SETTINGS_ORDER)
        n += set_flag(settings, 'doNotHyphenateCaps', SETTINGS_ORDER)
        # a stored "spelling: clean" makes Word skip checking text that has changed since
        for tag in ('embedTrueTypeFonts', 'embedSystemFonts', 'saveSubsetFonts', 'proofState'):
            for el in settings.findall(q('w:' + tag)):
                remove(el)
                n += 1
    n += apply_scheme(d, st, scheme)
    n += theme_font(d)
    n += clean_fonts(d)
    return n


def scheme_styles(d, st):
    """--scheme: {style id: colour} for the [text] table; names resolve to ids."""
    names = {wv(s.find(q('w:name'))): sid for sid, s in st.by_id.items()}
    out = {}
    for key, color in d.opts.scheme.get('text', {}).items():
        sid = key if key in st.by_id else names.get(key)
        if sid is None:
            if key not in d.findings['Colour scheme: style not found']:
                d.findings['Colour scheme: style not found'].append(key)
            continue
        out[sid] = color
    return out


def apply_scheme(d, st, text):
    """--scheme FILE: text colours of styles ([text]) and the fill of shaded table header rows
    ([table] header_fill). Applied after the house style made headings black."""
    n = 0
    for sid, color in text.items():
        el, created = ensure(style_el(st, sid, 'rPr'), 'color', RPR_ORDER)
        n += created + set_attrs(el, {'val': color})
    fill = d.opts.scheme.get('table', {}).get('header_fill')
    if fill:
        for tbl in d.xml(d.main).iter(TBL):
            rows = tbl.findall(q('w:tr'))
            for i, tr in enumerate(rows):
                if i and tr.find('w:trPr/w:tblHeader', NS) is None:
                    continue
                for shd in tr.iter(q('w:shd')):
                    if shd.getparent().tag == q('w:tcPr') and (shd.get(q('w:fill')) or 'auto') not in ('auto', 'FFFFFF') \
                            and shd.get(q('w:fill')) != fill:
                        shd.set(q('w:fill'), fill)
                        for k in ('themeFill', 'themeFillTint', 'themeFillShade'):
                            shd.attrib.pop(q('w:' + k), None)
                        n += 1
    return n


def theme_font(d):
    """--font NAME: the theme's heading and body fonts become NAME, so every style that uses the
    theme fonts (all of them, after 11.6) is set in it."""
    name = d.opts.font
    theme = d.xml((d.pkg.related(d.main, 'theme') or [None])[0])
    if not name or theme is None:
        return 0
    n = 0
    a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    for kind in ('majorFont', 'minorFont'):
        for latin in theme.iter(a + kind):
            el = latin.find(a + 'latin')
            if el is not None and (el.get('typeface') != name or el.get('panose') is not None):
                el.set('typeface', name)
                el.attrib.pop('panose', None)
                n += 1
    return n


def text_language(d):
    """The run language most characters carry (before 11.5 strips it)."""
    c = Counter()
    for p in d.xml(d.main).iter(P):
        for r, txt in text_runs(p):
            c[wv(r.find('w:rPr/w:lang', NS))] += len(txt)
    c.pop(None, None)
    return c.most_common(1)[0][0] if c else None


def clean_fonts(d):
    """Embedded font files and font-table entries nothing refers to any more."""
    n = 0
    ft_name = d.part('fontTable')
    ft = d.xml(ft_name)
    if ft is None:
        return 0
    for font in ft.findall(q('w:font')):
        for el in list(font):
            if local(el).startswith('embed'):
                remove(el)
                n += 1
    used = set()
    for name in d.pkg.names:
        if name.endswith('.xml') and name != ft_name and name.startswith(posixpath.dirname(d.main) + '/'):
            for rf in d.xml(name).iter(q('w:rFonts')):
                used |= {v for k, v in rf.attrib.items() if not k.endswith('Theme') and not k.endswith('theme')}
    theme = d.xml((d.pkg.related(d.main, 'theme') or [None])[0])
    if theme is not None:
        used |= {el.get('typeface') for el in theme.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}latin')}
    for font in ft.findall(q('w:font')):
        if wv(font, 'name') not in used and not KEEP_FONTS.match(wv(font, 'name') or ''):
            remove(font)
            n += 1
    for rel, target, rtype in d.pkg.rels(ft_name):
        if rtype.endswith('/font'):
            n += d.pkg.delete(target)
    return n


def rule_tables(d):
    """§11.7 Cells: 0.1" padding, top-left text; header paragraphs carry no spacing of their own."""
    st = d.styles()
    _, cell, head = roles(d, st)
    n = 0
    mar = lambda: [mk('w:' + side, {'w': CELL_MARGIN, 'type': 'dxa'}) for side in ('top', 'left', 'bottom', 'right')]
    want = [(f'top', CELL_MARGIN), ('left', CELL_MARGIN), ('bottom', CELL_MARGIN), ('right', CELL_MARGIN)]
    def set_mar(tblpr):
        cm = tblpr.find(q('w:tblCellMar'))
        if cm is not None and [(local(c), wv(c, 'w')) for c in cm] == want:
            return 0
        if cm is not None:
            remove(cm)
        cm = insert_ordered(tblpr, mk('w:tblCellMar'), TBLPR_ORDER)
        cm.extend(mar())
        return 1
    for root in d.stories:
        for tbl in root.iter(TBL):
            tp = tbl.find(q('w:tblPr'))
            if tp is None:
                tp = mk('w:tblPr')
                tbl.insert(0, tp)
            n += set_mar(tp)
            for tc in tbl.iter(q('w:tc')):
                tcp = tc.find(q('w:tcPr'))
                if tcp is None:
                    continue
                for el in tcp.findall(q('w:tcMar')) + tcp.findall(q('w:vAlign')):
                    if el.tag == q('w:vAlign') and wv(el) == 'top':
                        continue
                    remove(el)
                    n += 1
    used_tbl = {wv(el) for root in d.stories for el in root.iter(q('w:tblStyle'))}
    for sid in used_tbl:
        for s in st.chain(sid):
            if s.get(q('w:type')) != 'table':
                continue
            tp = s.find(q('w:tblPr'))
            if tp is not None and tp.find(q('w:tblCellMar')) is not None:
                n += set_mar(tp)
            for el in list(s.iter(q('w:vAlign'))):
                if wv(el) != 'top':
                    remove(el)
                    n += 1
    for sid in {cell, head} - {None}:
        n += set_attrs(ensure(style_el(st, sid, 'pPr'), 'jc', PPR_ORDER)[0], {'val': 'left'})
    if head in st.by_id:
        sp, _ = ensure(style_el(st, head, 'pPr'), 'spacing', PPR_ORDER)
        for k in ('before', 'after'):
            if sp.get(q('w:' + k)) != '0':
                sp.set(q('w:' + k), '0')
                n += 1
        for k in ('beforeAutospacing', 'afterAutospacing'):
            if sp.get(q('w:' + k)) is not None:
                del sp.attrib[q('w:' + k)]
                n += 1
    return n


def page_break_only(p):
    """A paragraph whose only content is page breaks (and empty runs)."""
    brs = [b for b in p.iter(q('w:br')) if wv(b, 'type') == 'page']
    return bool(brs) and is_empty_but(p, brs)


def is_empty_but(p, allowed):
    for el in tokens(p):
        if el in allowed:
            continue
        if el.tag == T and not (el.text or '').strip():
            continue
        if el.tag in LAYOUT_ONLY:
            continue
        return False
    return p.find('.//w:sectPr', NS) is None


def rule_heading_pages(d):
    """§11.8 Level-1 headings start a page through their style only: a manual page break right
    before one would leave a blank page, and an empty heading at the very end adds one."""
    st = d.styles()
    body = d.xml(d.main).find(q('w:body'))
    n = 0
    for p in list(body.iter(P)):
        if p.getparent() is not body or d.heading_level(p, st) != 1:
            continue
        prev = p.getprevious()
        while prev is not None and prev.tag != P and prev.tag != TBL and local(prev) in (
                'bookmarkStart', 'bookmarkEnd', 'proofErr', 'permStart', 'permEnd'):
            prev = prev.getprevious()
        if prev is None or prev.tag != P:
            continue
        text = Text(p).text.replace('\x00', '').strip()
        if page_break_only(prev) and not any(True for _ in prev.iter(q('w:bookmarkStart'))):
            remove(prev)
            d.log['Manual page breaks before level-1 headings removed'].append(snippet(text) or '(empty heading)')
            n += 1
            continue
        runs = [r for r in prev.findall(R) if Text(r).text.strip() or r.find(q('w:br')) is not None]
        if runs:
            last = runs[-1]
            brs = [b for b in last.findall(q('w:br')) if wv(b, 'type') == 'page']
            if brs and brs[-1] is [c for c in last if c.tag != RPR][-1]:
                remove(brs[-1])
                d.log['Manual page breaks before level-1 headings removed'].append(snippet(text))
                n += 1
    # empty headings at the end of the body (nothing but empty paragraphs after them)
    tail = [b for b in body if b.tag in (P, TBL)]
    while tail and tail[-1].tag == P and is_empty(tail[-1]) and tail[-1].find('.//w:sectPr', NS) is None:
        p = tail.pop()
        if d.heading_level(p, st) and not any(True for _ in p.iter(q('w:bookmarkStart'))):
            remove(p)
            d.log['Empty headings at the end of the document removed'].append(f'H{d.heading_level(p, st)} {p_style(p, st)}')
            n += 1
    return n


# ---------------------------------------------------------------------------------------------
# §9 TOC: defer page numbers to the layout engine.

def rule_toc(d):
    n = 0
    settings = d.settings
    if settings is not None:
        n += set_attrs(ensure(settings, 'updateFields', SETTINGS_ORDER)[0], {'val': 'true'})
    for root in d.stories:
        for el in root.iter(q('w:fldChar'), q('w:fldSimple')):
            if el.tag == q('w:fldChar') and wv(el, 'fldCharType') != 'begin':
                continue
            if wv(el, 'dirty') != 'true':
                el.set(q('w:dirty'), 'true')
                n += 1
    return n


RULES = [
    ('8.1 accept revisions', rule_accept_revisions),
    ('8.2 strip comments', rule_strip_comments),
    ('8.3 metadata / PII', rule_metadata),
    ('8.4 macros & embeds', rule_macros),
    ('8.5 hidden text', rule_hidden_text),
    ('8.6 proofing marks', rule_proofing),
    ('8.7 protection', rule_protection),
    ('8.9 temp bookmarks', rule_temp_bookmarks),
    ('8.10 rsid (optional)', rule_strip_rsid),
    ('11.1 unstyled body paragraphs', rule_unstyled),
    ('11.2 headings set as body text', rule_heading_gaps),
    ('11.3 captions and block titles', rule_captions),
    ('11.4 promote formatting into styles', rule_promote),
    ('11.5 strip direct formatting', rule_strip_direct),
    ('11.6 house style', rule_house_style),
    ('11.7 table cells', rule_tables),
    ('11.8 level-1 headings start a page', rule_heading_pages),
    ('11.9 captions stay with pictures', rule_figure_keep),
    ('1 page size & margins', rule_page_setup),
    ('2 table width', rule_table_width),
    ('3 table page breaks', rule_table_breaks),
    ('6 text sanitize', rule_text),
    ('7b blank lines', rule_blank_lines),
    ('7a empty headings', rule_empty_headings),
    ('10a lead-in "." before list', rule_leadin_dot),
    ('5b bold lead-in', rule_bold_leadin),
    ('4 list pagination', rule_list_pagination),
    ('8.8 orphans', rule_orphans),
    ('9 TOC refresh', rule_toc),
]


def rule_number(name):
    return name.split()[0]


def select_rules(only):
    """RULES limited to the given spec numbers (all of them when none are given)."""
    if not only:
        return RULES
    unknown = only - {rule_number(n) for n, _ in RULES}
    if unknown:
        sys.exit('✗ unknown rule number(s): ' + ', '.join(sorted(unknown)))
    return [(n, fn) for n, fn in RULES if rule_number(n) in only]


def converge(d, rules):
    """Full passes until one makes 0 changes. Returns [ {rule: count}, ... ] per pass."""
    passes = []
    while True:
        d.findings = defaultdict(list)
        counts = {name: fn(d) for name, fn in rules}
        passes.append(counts)
        if not any(counts.values()):
            return passes, True
        if len(passes) >= MAX_PASSES:
            return passes, False


# ---------------------------------------------------------------------------------------------
# Verification (Appendix E/F): package integrity on the result.

SAME_SLOT = {'footerReference': 'headerReference'}


def check_order(root):
    bad = 0
    for el in root.iter(*(q('w:' + k) for k in ORDERS)):
        order = ORDERS[local(el)]
        # headerReference/footerReference form one choice group and may interleave
        idx = [order.index(SAME_SLOT.get(local(c), local(c))) for c in el
               if local(c) in order and etree.QName(c).namespace == NS['w']]
        bad += idx != sorted(idx)
    return bad


def verify(d):
    results = []
    d.pkg.serialize()
    errors = []
    for name in d.pkg.names:
        if name.endswith(('.xml', '.rels')):
            try:
                etree.fromstring(d.pkg.raw[name], PARSER)
            except etree.XMLSyntaxError as e:
                errors.append(f'{name}: {e}')
    results.append(('every XML part parses', not errors, '; '.join(errors)))
    bad = sum(check_order(r) for r in d.word_parts)
    results.append(('pPr/trPr/tblPr/sectPr/settings children in schema order', bad == 0, f'{bad} out of order'))
    st = d.styles()
    missing = style_refs(d, st) - set(st.by_id)
    results.append(('every referenced style resolves', not missing, ', '.join(sorted(missing))))
    numbering = d.xml(d.part('numbering'))
    nums = {wv(x, 'numId'): wv(x.find(q('w:abstractNumId'))) for x in numbering.findall(q('w:num'))} \
        if numbering is not None else {}
    absts = {wv(a, 'abstractNumId') for a in numbering.findall(q('w:abstractNum'))} if numbering is not None else set()
    broken = [r for r in numbering_refs(d) if nums.get(r) not in absts]
    results.append(('every referenced numId resolves num -> abstractNum', not broken, ', '.join(broken)))
    dangling = [f'{rn}: {rel.get("Target")}' for rn in d.pkg.names if rn.endswith('.rels')
                for rel in d.pkg.xml(rn) if d.pkg.rel_target(rn, rel) not in (None, *d.pkg.raw)]
    results.append(('every internal relationship target exists', not dangling, '; '.join(dangling)))
    ct = d.xml(CONTENT_TYPES)
    overrides = {o.get('PartName')[1:] for o in ct.findall(q('ct:Override'))}
    defaults = {x.get('Extension').lower() for x in ct.findall(q('ct:Default'))}
    untyped = [n for n in d.pkg.names if n != CONTENT_TYPES and n not in overrides
               and n.rsplit('.', 1)[-1].lower() not in defaults]
    stale = [n for n in overrides if n not in d.pkg.raw]
    results.append(('content types cover exactly the parts', not untyped and not stale,
                    ', '.join(untyped + stale)))
    ids = [{wv(e, 'id') for r in d.stories for e in r.iter(q('w:' + t))}
           for t in ('commentRangeStart', 'commentRangeEnd', 'commentReference')]
    comments = d.xml(posixpath.join(posixpath.dirname(d.main), 'comments.xml'))
    ids.append({wv(c, 'id') for c in comments.findall(q('w:comment'))} if comments is not None else set())
    results.append(('comment id sets equal', all(s == ids[0] for s in ids), str([len(s) for s in ids])))
    return results


# ---------------------------------------------------------------------------------------------
# CLI.

def stage(title):
    print(f'\n\033[38;5;208m▶ {title}\033[0m')


def write_report(path, src, out, passes, converged, d, results, rules):
    lines = [f'# docx_fix report', '', f'- Input: `{src}`', f'- Output: `{out or "(check only)"}`',
             f'- Rules: {"all" if rules is RULES else ", ".join(rule_number(n) for n, _ in rules)}',
             f'- Passes: {len(passes)} ({"converged" if converged else "NOT converged"})', '',
             '## Changes per pass', '', '| Rule | ' + ' | '.join(f'Pass {i}' for i in range(1, len(passes) + 1)) + ' |',
             '|---|' + '---:|' * len(passes)]
    for name, _ in rules:
        lines.append(f'| {name} | ' + ' | '.join(str(p[name]) for p in passes) + ' |')
    lines += ['', '## Document stats', ''] + [f'- {k}: {v}' for k, v in d.stats.items()]
    lines += ['', '## Changes made', '']
    for title, items in d.log.items():
        lines += [f'### {title} ({len(items)})', ''] + [f'- {x}' for x in items] + ['']
    lines += ['## Findings for review', '']
    for title, items in d.findings.items():
        lines += [f'### {title} ({len(items)})', ''] + [f'- {x}' for x in items] + ['']
    lines += ['## Verification', ''] + [f'- {"PASS" if ok else "FAIL"} {name}' + (f' — {detail}' if not ok else '')
                                         for name, ok, detail in results]
    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')


def parse_args(argv):
    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument('mode', choices=['check', 'fix'])
    ap.add_argument('src')
    ap.add_argument('out', nargs='?')
    ap.add_argument('--heading-style', action='append', default=[])
    ap.add_argument('--leadin', action='append', default=[])
    ap.add_argument('--strip-rsid', action='store_true')
    ap.add_argument('--only', action='append', default=[])
    ap.add_argument('--font')
    ap.add_argument('--scheme', dest='scheme_file')
    try:
        args = ap.parse_args(argv)
        args.heading_styles = {k: int(v) for k, v in (x.split('=') for x in args.heading_style)}
        args.scheme = {}
        if args.scheme_file:
            with open(args.scheme_file, 'rb') as f:
                args.scheme = tomllib.load(f)
        colours = list(args.scheme.get('text', {}).values()) + list(args.scheme.get('table', {}).values())
        if not all(re.fullmatch(r'[0-9A-Fa-f]{6}|auto', str(v)) for v in colours):
            raise ValueError
    except (SystemExit, ValueError):
        print(__doc__, file=sys.stderr)
        sys.exit(2)
    args.leadins = set(args.leadin)
    args.only = {x.strip() for v in args.only for x in v.split(',') if x.strip()}
    if args.mode == 'check' and args.out:
        print(__doc__, file=sys.stderr)
        sys.exit(2)
    return args


def main():
    args = parse_args(sys.argv[1:])
    src = os.path.abspath(args.src)
    out = None
    if args.mode == 'fix':
        out = os.path.abspath(args.out or os.path.splitext(src)[0] + '_fixed.docx')
        if out == src:
            sys.exit('✗ refusing to overwrite the input file')

    stage('Read')
    pkg = Package(src)
    d = Doc(pkg, args)
    print(f'{os.path.basename(src)}: {len(pkg.names)} parts, main part {d.main}')
    rules = select_rules(args.only)

    stage('Rules (convergent passes)')
    passes, converged = converge(d, rules)
    width = max(len(n) for n, _ in rules)
    for name, _ in rules:
        counts = [p[name] for p in passes]
        if any(counts):
            print(f'  {name:<{width}}  ' + ' → '.join(map(str, counts)))
    print(f'  {len(passes)} passes, {"converged" if converged else "NOT converged"}')
    first_pass = sum(passes[0].values())

    stage('Findings')
    for title, items in list(d.log.items()) + list(d.findings.items()):
        print(f'  {title}: {len(items)}')

    stage('Verify')
    results = verify(d)
    for name, ok, detail in results:
        print(f'  {"✓" if ok else "✗"} {name}' + (f' — {detail}' if not ok else ''))
    ok = converged and all(r[1] for r in results)

    report = (os.path.splitext(out)[0] if out else os.path.splitext(src)[0] + '_check') + '.report.md'
    if out and ok:
        stage('Write')
        pkg.save(out)
        with zipfile.ZipFile(out) as z:
            bad = z.testzip()
        ok = bad is None
        print(f'  {out} ({os.path.getsize(src):,} → {os.path.getsize(out):,} bytes, {len(pkg.names)} parts)'
              + ('' if ok else f' — zip test failed on {bad}'))
    write_report(report, src, out if ok else None, passes, converged, d, results, rules)
    print(f'  report: {report}')

    if not ok:
        print('\n✗ finished with errors')
        sys.exit(1)
    if args.mode == 'check':
        print(f'\n{"✗" if first_pass else "✓"} check complete: {first_pass} change(s) needed')
        sys.exit(1 if first_pass else 0)
    print('\n✓ fix complete')


if __name__ == '__main__':
    main()

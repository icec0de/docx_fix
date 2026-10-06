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

See spec.md.
"""
import argparse
import copy
import hashlib
import os
import posixpath
import re
import sys
import zipfile
from collections import defaultdict
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
                if len(run) >= 2:
                    keep = [p for p in run if protected(p)] or run[:1]
                    for p in run:
                        if p not in keep:
                            remove(p)
                            n += 1
                run = []
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


def converge(d):
    """Full passes until one makes 0 changes. Returns [ {rule: count}, ... ] per pass."""
    passes = []
    while True:
        d.findings = defaultdict(list)
        counts = {name: fn(d) for name, fn in RULES}
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


def write_report(path, src, out, passes, converged, d, results):
    lines = [f'# docx_fix report', '', f'- Input: `{src}`', f'- Output: `{out or "(check only)"}`',
             f'- Passes: {len(passes)} ({"converged" if converged else "NOT converged"})', '',
             '## Changes per pass', '', '| Rule | ' + ' | '.join(f'Pass {i}' for i in range(1, len(passes) + 1)) + ' |',
             '|---|' + '---:|' * len(passes)]
    for name, _ in RULES:
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
    try:
        args = ap.parse_args(argv)
        args.heading_styles = {k: int(v) for k, v in (x.split('=') for x in args.heading_style)}
    except (SystemExit, ValueError):
        print(__doc__, file=sys.stderr)
        sys.exit(2)
    args.leadins = set(args.leadin)
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

    stage('Rules (convergent passes)')
    passes, converged = converge(d)
    width = max(len(n) for n, _ in RULES)
    for name, _ in RULES:
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
    write_report(report, src, out if ok else None, passes, converged, d, results)
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

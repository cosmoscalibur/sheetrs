"""Makes the fixtures of `tests/replace.rs`: `input.{xlsx,ods}` and, for each case of `CASES`,
`expected_<case>.{xlsx,ods}`, the workbook after LibreOffice itself replaced the text
(`Sheet.replaceAll`), plus `cases.json`.

Run it with a LibreOffice listening on a socket (26.8 until another is the reference):

    /opt/libreoffice26.8/program/soffice --headless --accept="socket,host=localhost,port=2002;urp;" &
    /opt/libreoffice26.8/program/python make_fixtures.py
"""
import json, os, sys, time
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import Locale
from com.sun.star.table import CellAddress
from com.sun.star.sheet import TableFilterField2
from com.sun.star.sheet.FilterOperator2 import EQUAL

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "..", "scripts"))
import fixture_determinism  # the same bytes on every run, see its docstring

local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def new():
    return desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))


OUT=os.path.dirname(os.path.abspath(__file__))
os.makedirs(OUT,exist_ok=True)
def build():
    doc=new(); sheets=doc.Sheets
    sh=sheets.getByIndex(0); sh.Name='T'
    sheets.insertNewByName('U Sheet',1)
    u=sheets.getByName('U Sheet')
    fm=doc.NumberFormats; loc=Locale('en','US','')
    def fmt(c,code):
        k=fm.queryKey(code,loc,False)
        if k==-1:k=fm.addNew(code,loc)
        c.NumberFormat=k
    def cell(r,c=0,s=sh): return s.getCellByPosition(c,r)
    cell(0).String='Hello World hello'
    cell(1).String='123'
    cell(2).Value=1234.5; fmt(cell(2),'#,##0.00')
    cell(3).Value=45000; fmt(cell(3),'yyyy-mm-dd')
    cell(4).Value=0.256; fmt(cell(4),'0.0%')
    cell(5).Formula='=SUM(B1;B2)'
    cell(6).Formula="=$'U Sheet'.A1+1"
    cell(7).Formula='="abc"&"d"'
    cell(8).String='abc'
    cell(9).Formula='=B1*2'
    cell(10).String='a1 b1 A1'
    cell(11).Formula='=IF(B1>0;"yes";"no")'
    cell(12).Value=7
    # What a cell shows is not what its input line says (`Find.formatted`, `SearchFormatted`).
    cell(25,7).Value=45000; fmt(cell(25,7),'D MMM YYYY')
    cell(26,7).Value=10; fmt(cell(26,7),'[$$-409]#,##0.00')
    cell(27,7).Value=5; fmt(cell(27,7),'"x"0')
    cell(28,7).Value=1; fmt(cell(28,7),'"yes";"yes";"no"')
    cell(29,7).Value=0.5; fmt(cell(29,7),'0.00%')
    cell(0,1).Value=1; cell(1,1).Value=2
    for c in (4,5,6): cell(13,c).String='repeat'
    for r in (19,20,21): cell(r).String='rowrep'
    cell(0,0,u).Value=5; cell(1,0,u).Value=7; cell(2,0,u).String='Hello U'
    # Notes: one paragraph, several (with an empty one), on a cell with nothing else, on another sheet.
    def note(s,col,row,text):
        a=CellAddress(); a.Sheet=s.RangeAddress.Sheet; a.Column=col; a.Row=row
        s.Annotations.insertNew(a,text)
    note(sh,0,0,'Hello note\nsecond hello paragraph')
    note(sh,1,0,'HELLO again\n\nlast hello')
    note(sh,3,24,'Hello on an empty cell')
    note(u,0,0,'Hello U note')
    return doc
def build_filtered():
    """A workbook with an autofilter (anonymous database range A1:B6, key = x) that hides rows 3 and
    5, and a row hidden by hand under it (`Find.filtered`, `SearchFiltered`). Kept apart from
    `input.*`: the editor does not adjust a filter yet, so a workbook that has one cannot take the
    structure operations the other tests run on `input.*`."""
    doc=new(); f=doc.Sheets.getByIndex(0); f.Name='F'
    def cell(r,c=0): return f.getCellByPosition(c,r)
    cell(0,0).String='key'; cell(0,1).String='note'
    for i,(k,n) in enumerate([('x','filt a'),('y','filt b'),('x','filt c'),('y','filt d'),('x','filt e')]):
        cell(i+1,0).String=k; cell(i+1,1).String=n
    cell(8,1).String='filt visible'; cell(9,1).String='filt by hand'
    f.Rows.getByIndex(9).IsVisible=False
    ctrl=doc.getCurrentController()
    rng=f.getCellRangeByName('A1:B6'); ctrl.select(rng)
    disp=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper",ctx)
    disp.executeDispatch(ctrl.getFrame(),".uno:DataFilterAutoFilter","",0,())
    time.sleep(0.5)
    fd=rng.createFilterDescriptor(True)
    ff=TableFilterField2(); ff.Field=0; ff.IsNumeric=False; ff.StringValue='x'; ff.Operator=EQUAL
    fd.setFilterFields2((ff,))
    rng.filter(fd)
    return doc
# What is typed into a cell takes the number format LibreOffice gives that text when the cell is in
# General (`typed_general`, one text per row of column A of `V`) and keeps the cell's own format
# otherwise, except for a boolean (`typed_formats`: column B, one `(format, text)` per row).
TYPED_GENERAL=['50%','-50%','1.5%','1,234.5%','03/15/2024','1/2','1/2/24','Mar 5','Mar 15 2024','15-Mar-24','March 15, 2024','Jan 2024',
 '2024-03-15','03/15/2024 10:20','03/15/2024 10:20 PM','03/15/2024 10:20:30','2024-03-15 10:20','2024-03-15 10:20:30','2024-03-15T10:20',
 '12:30','12:30:45','12:30:45.5','10:20 PM','12:30 am','1:2','25:00','48:30','true','TRUE','false','1E3','1e-3','5E2','-1E3','12.5E-2',
 '$12.50','$1,234.50','($5)','-$5','5$','$ 5','1 1/2','-1 1/2','2 3/4','0 1/2','12.50 USD','15 Mar 2024','2024/03/15','13/13/2013',
 '12','12.5','1,234.5','-5','(5)','.5','+5','00123','abc','=1+1','5 EUR','1.2.2024','12,5','5-3']
TYPED_FORMATS=[('0.00','50%'),('0.00','2024-03-15'),('0.00','1E3'),('0.00%','12'),('yyyy-mm-dd','12'),('yyyy-mm-dd','12:30'),
 ('@','50%'),('@','true'),('[$$-409]#,##0.00','true'),('0.00','true'),('0.00','false'),('hh:mm:ss','1/2'),('"x"0','12'),('0%','$5'),('#,##0.00','abc')]
def build_typed():
    doc=new(); v=doc.Sheets.getByIndex(0); v.Name='V'
    fm=doc.NumberFormats; loc=Locale('en','US','')
    for i,t in enumerate(TYPED_GENERAL): v.getCellByPosition(0,i).String='T:'+t
    for i,(code,t) in enumerate(TYPED_FORMATS):
        c=v.getCellByPosition(1,i)
        if code!='General':
            k=fm.queryKey(code,loc,False)
            c.NumberFormat=k if k!=-1 else fm.addNew(code,loc)
        c.String='T:'+t
    return doc
def save(doc,stem):
    doc.storeToURL('file://'+OUT+'/'+stem+'.xlsx',(pv('FilterName','Calc MS Excel 2007 XML'),))
    doc.storeToURL('file://'+OUT+'/'+stem+'.ods',(pv('FilterName','calc8'),))
    fixture_determinism.normalize(OUT+'/'+stem+'.xlsx'); fixture_determinism.normalize(OUT+'/'+stem+'.ods')
ONLY=set(sys.argv[1:])  # the cases to make again; none: every one, with their inputs
if not ONLY:
    doc=build(); save(doc,'input'); doc.close(True)
    doc=build_filtered(); save(doc,'input_filtered'); doc.close(True)
if not ONLY or 'typed_general' in ONLY:
    doc=build_typed(); save(doc,'input_typed'); doc.close(True)
CASES=[
 dict(name='text_ci',text='hello',with_='bye'),
 dict(name='match_case',text='hello',with_='bye',match_case=True),
 dict(name='entire',text='abc',with_='xyz',entire_cells=True),
 dict(name='number_digits',text='34',with_='99'),
 dict(name='number_to_text',text='1234',with_='x'),
 dict(name='text_to_number',text='abc',with_='12',entire_cells=True),
 dict(name='text_to_bool',text='abc',with_='true',entire_cells=True),
 dict(name='text_to_formula',text='abc',with_='=1+1',entire_cells=True),
 dict(name='clear',text='abc',with_='',entire_cells=True),
 dict(name='date_year',text='2023',with_='2024'),
 dict(name='percent',text='25',with_='30'),
 dict(name='formula_ref',text='B1',with_='B2',search_in='formula'),
 dict(name='formula_sheet_ref',text='A1',with_='A2',search_in='formula',case=True),
 dict(name='formula_string_literal',text='abc',with_='xyz',search_in='formula'),
 dict(name='formula_function',text='SUM',with_='AVERAGE',search_in='formula'),
 dict(name='formula_to_text',text='=',with_='',search_in='formula',entire_cells=False),
 dict(name='value_include_formulas',text='abcd',with_='zz',include_formulas=True),
 dict(name='value_excludes_formulas',text='abcd',with_='zz'),
 dict(name='formatted_number',text='1,234',with_='9,999',formatted=True),
 dict(name='formatted_number_off',text='1,234',with_='9,999'),
 dict(name='formatted_date',text='Mar',with_='Apr',formatted=True),
 dict(name='formatted_currency',text='$10',with_='$20',formatted=True),
 dict(name='formatted_custom',text='x5',with_='x9',formatted=True),
 dict(name='formatted_label',text='yes',with_='no',formatted=True),
 dict(name='formatted_percent',text='25.6',with_='30',formatted=True),
 dict(name='formatted_percent_precision',text='50.00',with_='75',formatted=True),
 dict(name='formatted_entire',text='1,234.50',with_='5',formatted=True,entire_cells=True),
 dict(name='formatted_regex',text='(\\d+),(\\d+)',with_='$2-$1',mode='regex',formatted=True),
 dict(name='formatted_formula_search',text='SUM',with_='AVERAGE',search_in='formula',formatted=True),
 dict(name='regex_groups',text='(l+)o',with_='[$1|&]',mode='regex'),
 dict(name='wildcards',text='h*o',with_='X',mode='wildcards',entire_cells=True),
 dict(name='range_scope',text='hello',with_='bye',scope=[{'range':'T!A1:A2'}]),
 dict(name='repeated_cells',text='repeat',with_='again'),
 dict(name='repeated_cells_partial',text='repeat',with_='again',scope=[{'range':'T!F14'}]),
 dict(name='repeated_rows',text='rowrep',with_='row'),
 dict(name='repeated_rows_partial',text='rowrep',with_='row',scope=[{'range':'T!A21'}]),
 dict(name='sheet_scope',text='Hello',with_='Bye',scope=[{'sheet':'U Sheet'}]),
 dict(name='comment_ci',text='hello',with_='bye',search_in='comment'),
 dict(name='comment_match_case',text='hello',with_='bye',search_in='comment',match_case=True),
 dict(name='comment_entire',text='Hello U note',with_='x y',search_in='comment',entire_cells=True),
 dict(name='comment_regex',text='(h)ello',with_='[$1|&]',search_in='comment',mode='regex'),
 dict(name='comment_wildcards',text='*empty*',with_='nothing',search_in='comment',mode='wildcards',entire_cells=True),
 dict(name='comment_range_scope',text='hello',with_='bye',search_in='comment',scope=[{'range':'T!A1:A1'}]),
 dict(name='comment_sheet_scope',text='Hello',with_='Bye',search_in='comment',scope=[{'sheet':'U Sheet'}]),
 dict(name='filtered_left_out',text='filt',with_='done',input='input_filtered'),
 dict(name='filtered_included',text='filt',with_='done',filtered=True,input='input_filtered'),
 # What it makes of a text typed as a formula that is not one, or that names what it does not have.
 dict(name='invalid_open_paren',text='*2',with_='*((',search_in='formula'),
 dict(name='invalid_unknown_function',text='SUM',with_='SUMX',search_in='formula'),
 dict(name='invalid_missing_sheet',text='U Sheet',with_='Nope',search_in='formula'),
 dict(name='invalid_trailing_operator',text='B1*2',with_='B1*',search_in='formula'),
 dict(name='invalid_two_operands',text='B1*2',with_='B1 2',search_in='formula'),
 dict(name='invalid_extra_close',text='B1*2',with_='B1*2)',search_in='formula'),
 dict(name='invalid_function_left_open',text='B1,B2)',with_='B1,B2',search_in='formula'),
 dict(name='invalid_unknown_name',text='B1*2',with_='B1*Total',search_in='formula'),
 dict(name='invalid_error_constant',text='B1*2',with_='#REF!+1',search_in='formula'),
 dict(name='invalid_empty_paren',text='B1*2',with_='(',search_in='formula'),

 # The number format of what is typed (`formats`: the formats of the cells are compared too).
 dict(name='typed_general',text='T:(.*)',with_='$1',mode='regex',scope=[{'range':'V!A1:A%d'%len(TYPED_GENERAL)}],input='input_typed',formats=True),
 dict(name='typed_formats',text='T:(.*)',with_='$1',mode='regex',scope=[{'range':'V!B1:B%d'%len(TYPED_FORMATS)}],input='input_typed',formats=True),
]
for c in CASES:
    if ONLY and c['name'] not in ONLY: continue
    doc={'input_filtered':build_filtered,'input_typed':build_typed}.get(c.get('input'),build)()
    scope=c.get('scope')
    targets=[]
    for s in (scope or [{'workbook':True}]):
        if 'range' in s:
            sn,rng=s['range'].split('!'); targets.append(doc.Sheets.getByName(sn).getCellRangeByName(rng))
        elif 'sheet' in s: targets.append(doc.Sheets.getByName(s['sheet']))
        else: targets+= [doc.Sheets.getByIndex(i) for i in range(doc.Sheets.Count)]
    for t in targets:
        d=t.createReplaceDescriptor() if hasattr(t,'createReplaceDescriptor') else None
        d.SearchString=c['text']; d.ReplaceString=c['with_']
        d.SearchType={'formula':0,'comment':2}.get(c.get('search_in'),1)
        d.SearchCaseSensitive=bool(c.get('match_case'))
        d.SearchWords=bool(c.get('entire_cells'))
        d.setPropertyValue('SearchFormatted',bool(c.get('formatted')))
        d.setPropertyValue('SearchFiltered',bool(c.get('filtered')))
        d.SearchRegularExpression=c.get('mode')=='regex'
        d.SearchWildcard=c.get('mode')=='wildcards'
        t.replaceAll(d)
    save(doc,'expected_'+c['name']); doc.close(True)
json.dump(CASES,open(OUT+'/cases.json','w'),indent=1)
print(len(CASES))

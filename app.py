import io
import html
import re
import zipfile
import traceback
from urllib.parse import urlparse

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title='Thermo Fisher | URL Tools', page_icon='🌐', layout='wide')
st.markdown('''<style>
.block-container {padding-top:2rem;padding-bottom:3rem;max-width:1500px;}
.hero {padding:2rem 2.2rem;border-radius:24px;background:linear-gradient(135deg,#071a2c,#0b5d5b);color:white;margin-bottom:1.2rem;box-shadow:0 12px 35px rgba(0,0,0,.14)}
.hero h1 {font-size:2.35rem;margin:0 0 .35rem 0;}
.hero p {opacity:.88;margin:0;font-size:1.05rem;}
.pill {display:inline-block;padding:.32rem .68rem;margin:.18rem .18rem .18rem 0;border-radius:999px;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);font-size:.88rem;}
[data-testid="stMetric"] {background:rgba(127,127,127,.07);border:1px solid rgba(127,127,127,.15);padding:14px;border-radius:18px;}
</style><div class="hero"><h1>🌐 Thermo Fisher | URL Tools</h1><p>Marketing URL conversion, Product Inclusion Lists and AEM translation URL generation in one place.</p><br><span class="pill">URL Converter</span><span class="pill">AEM URLs</span><span class="pill">Naming Converter</span><span class="pill">Excel exports</span></div>''', unsafe_allow_html=True)

AUTHOR = 'https://author-prod-use1.aemprod.thermofisher.net'
LOCALES = {
    'de-DE': {'flag':'🇩🇪','label':'German / Germany','root':'europe','local':'de-de','regional':'en-de'},
    'es-ES': {'flag':'🇪🇸','label':'Spanish / Spain','root':'europe','local':'es-es','regional':'en-es'},
    'fr-FR': {'flag':'🇫🇷','label':'French / France','root':'europe','local':'fr-fr','regional':'en-fr'},
    'ja-JP': {'flag':'🇯🇵','label':'Japanese / Japan','root':'japan','local':'ja-jp','regional':'en-jp'},
    'ko-KR': {'flag':'🇰🇷','label':'Korean / Korea','root':'ipac','local':'ko-kr','regional':'en-kr'},
    'zh-CN': {'flag':'🇨🇳','label':'Chinese / Mainland China','root':'greater-china','local':'zh-cn','regional':'en-cn'},
    'pt-BR': {'flag':'🇧🇷','label':'Portuguese / Brazil','root':'latin-america','local':'pt-br','regional':'en-br'},
    'zh-TW': {'flag':'🇹🇼','label':'Traditional Chinese / Taiwan','root':'ipac','local':'zh-tw','regional':'en-tw'},
    'es-LATAM': {'flag':'🌎','label':'Spanish / Latin America','special':True},
}
LATAM = {
    'Argentina': {'flag':'🇦🇷','local':'es-ar','regional':'en-ar'},
    'Chile': {'flag':'🇨🇱','local':'es-cl','regional':'en-cl'},
    'Mexico': {'flag':'🇲🇽','local':'es-mx','regional':'en-mx'},
}
LANGUAGE_MAP = {lang: f"/content/lifetech/{cfg['root']}/{cfg['regional']}" for lang,cfg in LOCALES.items() if not cfg.get('special')}
LANGUAGE_MAP['es-LATAM'] = '/content/lifetech/latin-america/en-mx'

# URL Converter: original multi-sheet workflow.
def get_default_project_code():
    return 'GTS2600XX'


def detect_header_marketing(df):
    for i, row in df.iterrows():
        cells = [str(cell).lower() if pd.notna(cell) else '' for cell in row]
        if any('page title' in cell for cell in cells) and any('url' in cell for cell in cells):
            return i
    return 0


def detect_header_language(df):
    for i,row in df.iterrows():
        if any(re.match(r'[a-z]{2}-[A-Z]{2}',str(c)) for c in row if pd.notna(c)):
            return i
    return 0


def normalize_lang_column(colname):
    match = re.match(r'([a-z]{2}-[A-Z]{2})',str(colname))
    return match.group(1) if match else None


def make_excel_buffer(df):
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine='openpyxl') as writer:
        df.to_excel(writer,index=False)
    return buf.getvalue()


def make_product_excel_buffer(df_ids):
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine='openpyxl') as writer:
        ids=df_ids.copy()
        ids.columns=['ERP']
        ids.to_excel(writer,index=False,startrow=1,header=False,sheet_name='Sheet1')
        ws=writer.sheets['Sheet1']
        ws['A1']='ERP'
        for cell in ws['A']:
            cell.number_format='@'
            cell.alignment=Alignment(horizontal='left')
        ws.column_dimensions[get_column_letter(1)].width=20
    return buf.getvalue()


def process_all_sheets(file):
    xls=pd.ExcelFile(file)
    marketing_records=[]
    product_records=[]
    for idx,sheet in enumerate(xls.sheet_names):
        preview=pd.read_excel(xls,sheet_name=sheet,header=None)
        header_row=detect_header_marketing(preview) if idx==0 else detect_header_language(preview)
        df=pd.read_excel(xls,sheet_name=sheet,header=header_row)
        df.columns=[str(c).strip().replace('\n',' ') for c in df.columns]
        lang_cols={col:normalize_lang_column(col) for col in df.columns if normalize_lang_column(col) in LANGUAGE_MAP}
        if not lang_cols:
            continue
        for _,row in df.iterrows():
            cells=list(row)
            pid=None
            for cell in cells:
                if isinstance(cell,str):
                    m=re.search(r'A\d{3,12}',cell)
                    if m:
                        pid=m.group(0)
                        break
                    if '/product/' in cell:
                        seg=cell.split('/product/',1)[1]
                        pid=re.split(r'[/\?#]',seg)[0]
                        break
            murl=next((c for c in cells if isinstance(c,str) and '/home/' in urlparse(c).path),None)
            if pid:
                for col,lang in lang_cols.items():
                    val=row.get(col)
                    if pd.notna(val) and str(val).strip().lower() in ['x','yes','✓','✔']:
                        product_records.append({'Product ID':pid,'Language':lang})
            elif murl:
                try:
                    path=urlparse(murl).path
                    cleaned='/home/'+re.sub(r'.*?/home/','',path)
                    cleaned=re.sub(r'\.html($|[\?#])',r'\1',cleaned)
                except (ValueError,TypeError):
                    continue
                for col,lang in lang_cols.items():
                    val=row.get(col)
                    if pd.notna(val) and str(val).strip().lower() in ['x','yes','✓','✔']:
                        marketing_records.append({'Original URL':murl,'Language':lang,'Localized Path':LANGUAGE_MAP[lang]+cleaned})
    return pd.DataFrame(marketing_records),pd.DataFrame(product_records)


def render_url_converter():
    st.subheader('🔗 URL Converter')
    st.caption('Process all sheets to generate Marketing URLs and Product Inclusion Lists.')
    project_code=st.text_input('GTS ID',value=get_default_project_code(),key='converter_gts').strip() or get_default_project_code()
    st.session_state['shared_gts']=project_code
    uploaded=st.file_uploader('Upload an Excel File',type=['xlsx'],key='converter_upload')
    if not uploaded:
        if st.button('🔁 Reset',key='converter_reset_empty'):
            reset_application()
        for k in ('shared_languages','shared_marketing','shared_product'):
            st.session_state.pop(k,None)
        st.info('Upload an .xlsx workbook to begin.')
        return
    try:
        df_marketing,df_product=process_all_sheets(io.BytesIO(uploaded.getvalue()))
    except Exception as exc:
        st.error(f'Could not process workbook: {exc}')
        with st.expander('Technical error details'):
            st.code(traceback.format_exc())
        return
    st.session_state['shared_languages']=sorted(set(df_marketing.get('Language',pd.Series(dtype=str)).tolist()+df_product.get('Language',pd.Series(dtype=str)).tolist()))
    st.session_state['shared_marketing']=not df_marketing.empty
    st.session_state['shared_product']=not df_product.empty
    buffers_all={}
    if not df_marketing.empty:
        df_marketing=df_marketing.sort_values(by=['Language','Original URL'],ascending=[True,True]).reset_index(drop=True)
        st.header('Marketing URLs')
        st.dataframe(df_marketing,use_container_width=True)
        name=f'{project_code} - Converted URLs.xlsx'
        content=make_excel_buffer(df_marketing)
        st.download_button('📥 Download Converted URLs',content,file_name=name,mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key='converter_marketing')
        buffers_all[name]=content
    else:
        st.warning('No marketing URLs found.')
    if not df_product.empty:
        st.header('Product Inclusion List')
        for lang in sorted(df_product['Language'].unique()):
            df_lang=df_product[df_product['Language']==lang][['Product ID']]
            st.subheader(lang)
            st.table(df_lang)
            content=make_product_excel_buffer(df_lang)
            name=f'Product Inclusion List_{project_code}_{lang}.xlsx'
            st.download_button('📥 Download Inclusion List',content,file_name=name,mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key=f'converter_product_{lang}')
            buffers_all[name]=content
    else:
        st.warning('No product entries found.')
    if buffers_all:
        zip_buf=io.BytesIO()
        with zipfile.ZipFile(zip_buf,'w',compression=zipfile.ZIP_DEFLATED) as zf:
            for name,content in buffers_all.items():
                zf.writestr(name,content)
        _,center,right=st.columns([1,1,1])
        with center:
            st.download_button('📥 Download All',zip_buf.getvalue(),file_name=f'{project_code} - All Downloads.zip',mime='application/zip',key='converter_all')
        with right:
            if st.button('🔁 Reset',key='converter_reset'):
                reset_application()

# AEM URLs: original first-sheet translation workflow.
def clean_header(v):
    return str(v or '').split('\n')[0].strip()


def page_path(url):
    p=urlparse(str(url).strip()).path
    m=re.search(r'/home/(.+)$',p)
    if not m:
        return None
    return 'home/'+m.group(1)


def urls_for(path,cfg):
    root,local,regional=cfg['root'],cfg['local'],cfg['regional']
    nohtml=path[:-5] if path.endswith('.html') else path
    return {
        'Local Author':f'{AUTHOR}/editor.html/content/lifetech/{root}/{local}/{path}',
        'Regional Author':f'{AUTHOR}/editor.html/content/lifetech/{root}/{regional}/{path}',
        'Regional Sites / Rollout':f'{AUTHOR}/sites.html/content/lifetech/{root}/{regional}/{nohtml}',
        'Global Author':f'{AUTHOR}/editor.html/content/lifetech/global/en/{path}',
    }


def parse(file_bytes):
    wb=load_workbook(io.BytesIO(file_bytes),data_only=False)
    ws=wb[wb.sheetnames[0]]
    header_row=None
    url_col=None
    for r in range(1,min(ws.max_row,20)+1):
        for c in range(1,ws.max_column+1):
            if 'URL in AEM' in str(ws.cell(r,c).value or ''):
                header_row,url_col=r,c
                break
        if header_row:
            break
    if not header_row:
        raise ValueError("Could not find the 'URL in AEM' header in the first sheet.")
    lang_cols={}
    for c in range(1,ws.max_column+1):
        h=clean_header(ws.cell(header_row,c).value)
        if h in LOCALES:
            lang_cols[h]=c
    rows=[]
    for r in range(header_row+1,ws.max_row+1):
        url=ws.cell(r,url_col).value
        if not url or not str(url).strip().startswith(('http://','https://')):
            continue
        title=ws.cell(r,max(1,url_col-1)).value or ''
        selected=[lang for lang,c in lang_cols.items() if str(ws.cell(r,c).value or '').strip().lower()=='x']
        if selected:
            rows.append((r,str(title),str(url).strip(),selected))
    return ws.title,rows


def style_url_sheet(ws):
    headers=['Flag','Language','Page Title','Original URL','Local Author','Regional Author','Regional Sites / Rollout','Global Author']
    ws.append(headers)
    ws.sheet_view.showGridLines=False
    for cell in ws[1]:
        cell.fill=PatternFill('solid',fgColor='0B1F33')
        cell.font=Font(color='FFFFFF',bold=True)
        cell.alignment=Alignment(vertical='center',wrap_text=True)
    ws.row_dimensions[1].height=34
    for i,w in enumerate([8,28,34,58,70,70,70,70],1):
        ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes='A2'


def append_url_row(ws,flag,language,title,original,generated):
    ws.append([flag,language,title,original,generated['Local Author'],generated['Regional Author'],generated['Regional Sites / Rollout'],generated['Global Author']])
    r=ws.max_row
    thin=Side(style='thin',color='D8E2E8')
    for c in ws[r]:
        c.alignment=Alignment(vertical='top',wrap_text=True)
        c.border=Border(bottom=thin)
    ws.cell(r,1).alignment=Alignment(horizontal='center',vertical='top',wrap_text=True)
    for col in range(4,9):
        c=ws.cell(r,col)
        c.font=Font(color='008000',underline='single')
        c.hyperlink=c.value


def make_excel(rows,latam_countries):
    wb=Workbook()
    wb.remove(wb.active)
    detected=[]
    for _,_,_,langs in rows:
        for lang in langs:
            if lang not in detected:
                detected.append(lang)
    for lang in detected:
        ws=wb.create_sheet(lang[:31])
        style_url_sheet(ws)
        for _,title,original,langs in rows:
            if lang not in langs:
                continue
            path=page_path(original)
            if not path:
                continue
            if lang=='es-LATAM':
                for country in latam_countries:
                    x=LATAM[country]
                    cfg={'root':'latin-america','local':x['local'],'regional':x['regional']}
                    append_url_row(ws,x['flag'],f'es-LATAM — {country}',title,original,urls_for(path,cfg))
            else:
                cfg=LOCALES[lang]
                append_url_row(ws,cfg['flag'],f"{lang} — {cfg['label']}",title,original,urls_for(path,cfg))
        ws.auto_filter.ref=ws.dimensions
    info=wb.create_sheet('Locale Map')
    info.sheet_view.showGridLines=False
    info.append(['Language','Flag','Local root','Regional / Rollout root'])
    for lang,cfg in LOCALES.items():
        if lang=='es-LATAM':
            continue
        info.append([lang,cfg['flag'],f"/content/lifetech/{cfg['root']}/{cfg['local']}",f"/content/lifetech/{cfg['root']}/{cfg['regional']}"])
    for country,x in LATAM.items():
        info.append([f'es-LATAM — {country}',x['flag'],f"/content/lifetech/latin-america/{x['local']}",f"/content/lifetech/latin-america/{x['regional']}"])
    for c in info[1]:
        c.fill=PatternFill('solid',fgColor='0B6B68')
        c.font=Font(color='FFFFFF',bold=True)
    for col,w in enumerate([24,10,52,52],1):
        info.column_dimensions[get_column_letter(col)].width=w
    for row in info.iter_rows():
        for c in row:
            c.alignment=Alignment(vertical='top',wrap_text=True)
    info.freeze_panes='A2'
    out=io.BytesIO()
    wb.save(out)
    return out.getvalue()


def render_aem_urls():
    st.subheader('🌐 AEM Translation URL Builder')
    st.caption('First sheet only · Automatic locale detection · One worksheet per language · Wrapped Excel URLs')
    uploaded=st.file_uploader('Drop your translation request Excel here',type=['xlsx','xlsm'],help='Only the first worksheet is analyzed.',key='aem_upload')
    if not uploaded:
        st.info('Upload an .xlsx or .xlsm file to begin. The original workbook is not modified.')
        if st.button('🔁 Reset',key='aem_reset'):
            reset_application()
        return
    try:
        sheet,rows=parse(uploaded.getvalue())
    except Exception as exc:
        st.error(str(exc))
        return
    langs=sorted({lang for _,_,_,selected in rows for lang in selected})
    c1,c2,c3=st.columns(3)
    c1.metric('First sheet',sheet)
    c2.metric('Pages with selections',len(rows))
    c3.metric('Languages detected',len(langs))
    st.markdown('### Detected languages')
    st.markdown(' '.join(f"**{LOCALES[lang]['flag']} {lang}**" for lang in langs) if langs else 'No languages selected.')
    latam_countries=list(LATAM)
    if 'es-LATAM' in langs:
        st.info('es-LATAM is selected. Choose the Latin America destinations to generate.')
        latam_countries=st.multiselect('LATAM destinations',list(LATAM),default=list(LATAM),key='aem_latam')
    with st.expander('Preview detected pages',expanded=True):
        for _,title,url,selected in rows:
            st.markdown(f"**{title or 'Untitled page'}**  \n{url}  \n"+' · '.join(f"{LOCALES[lang]['flag']} {lang}" for lang in selected))
    try:
        output=make_excel(rows,latam_countries)
    except Exception as exc:
        st.error(f'Could not generate Excel: {exc}')
        return
    st.success('Your URL workbook is ready — each detected language has its own worksheet.')
    st.download_button('⬇️ Download generated Excel',output,file_name='aem_translation_urls.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True,key='aem_download')
    if st.button('🔁 Reset',key='aem_reset'):
        reset_application()


# Naming Converter: receives workbook metadata from Tab 1.
LANGUAGE_SUFFIXES = {
    'de-DE':'DE', 'es-ES':'ES', 'fr-FR':'FR', 'ja-JP':'JP',
    'ko-KR':'KR', 'zh-CN':'CN', 'zh-TW':'TW', 'pt-BR':'BR',
    'es-LATAM':'LATAM',
}


def reset_application():
    """Clear all app widget data and reload Streamlit (equivalent to a clean refresh)."""
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()


def parse_wordbee_text(raw):
    """Parse Wordbee table rows and flattened browser clipboard text."""
    text = html.unescape(raw).strip()
    if not text:
        return {'order': '', 'request': '', 'requester': ''}
    plain = re.sub(r'\[([^\]]+)\]\([^)]*\)', r'\1', text)
    plain = re.sub(r'<[^>]*>', ' ', plain)
    # Markdown table separators are not data.
    lines = [line.strip() for line in plain.splitlines()
             if line.strip() and not re.fullmatch(r'[|:\s-]+', line.strip())]
    plain = ' '.join(lines)
    plain = re.sub(r'\s+', ' ', plain).strip()
    fields = [re.sub(r'\*', '', x).strip() for x in re.split(r'\s*\|\s*|\t+', plain) if x.strip(' *|')]
    order_match = re.search(r'(?<!\d)\d{5,12}(?!\d)', plain)
    order = order_match.group(0) if order_match else ''
    request = fields[0] if fields else plain
    # Browser clipboard may flatten all columns into a single text string.
    if order:
        request = re.split(r'(?<!\d)' + re.escape(order) + r'(?!\d)', request, maxsplit=1)[0]
    request = re.sub(r'^Text Translation:\s*', '', request, flags=re.I).strip(' *|')
    requester = ''
    # Last column in Wordbee is the requester, including when a browser
    # copies the entire row with whitespace instead of tabs.
    if len(fields) >= 5:
        requester = fields[-1].strip(' *|')
    if not requester:
        tail = plain[order_match.end():] if order_match else ''
        tail = re.sub(r'(?i)View document\(s\)|Corporation|Thermo Fisher', ' ', tail)
        tail = re.sub(r'[|*]+', ' ', tail).strip()
        candidates = re.findall(r"[A-Za-zÀ-ÿ]+(?:[-'][A-Za-zÀ-ÿ]+)?\s+[A-Za-zÀ-ÿ]+(?:[-'][A-Za-zÀ-ÿ]+)?", tail)
        if candidates:
            requester = candidates[-1]
    return {'order': order, 'request': request, 'requester': requester}


def build_naming_results(order,project,requester,gts,languages,marketing,product):
    words = requester.strip().split()
    if len(words)<2:
        raise ValueError('Enter the requester first and last name.')
    person = words[0][0].upper()+''.join(words[1:])
    stem = f'{gts}_Web_{person}'
    base = f'{stem}_{project.strip()}'
    suffix = ('_AEM' if marketing else '')+('_PDP' if product else '')
    return {
        'Wordbee Name':f'GTS-{order}-{base}{suffix}',
        'AEM Names':[(lang,f'{base}_{LANGUAGE_SUFFIXES[lang]}') for lang in languages if lang in LANGUAGE_SUFFIXES],
        'Shared Folder Name':stem,
    }


def make_naming_excel(result, gts, order, project, requester, languages, marketing, product):
    """Generate a readable, wrapped, styled workbook from the confirmed conversion."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Naming Convention'
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = 'A5'
    ws.merge_cells('A1:C1')
    ws['A1'] = 'NAMING CONVENTION'
    ws['A1'].fill = PatternFill('solid', fgColor='0B1F33')
    ws['A1'].font = Font(name='Aptos Display', size=18, bold=True, color='FFFFFF')
    ws['A1'].alignment = Alignment(vertical='center', indent=1)
    ws.row_dimensions[1].height = 42
    ws.merge_cells('A2:C2')
    ws['A2'] = f'{gts}  |  Wordbee Order {order}'
    ws['A2'].font = Font(name='Aptos', size=11, color='0B6B68', bold=True)
    ws['A2'].alignment = Alignment(vertical='center', indent=1)
    ws.row_dimensions[2].height = 28
    for col, width in [('A', 24), ('B', 76), ('C', 24)]:
        ws.column_dimensions[col].width = width
    for cell, value in zip(ws[4], ['TYPE', 'GENERATED NAME', 'LANGUAGE']):
        cell.value = value
        cell.fill = PatternFill('solid', fgColor='0B6B68')
        cell.font = Font(name='Aptos', bold=True, color='FFFFFF')
        cell.alignment = Alignment(vertical='center', wrap_text=True)
    ws.row_dimensions[4].height = 28
    records = [('Wordbee Name', result['Wordbee Name'], 'All selected languages')]
    records += [('AEM Name', value, lang) for lang, value in result['AEM Names']]
    records.append(('Shared Folder Name', result['Shared Folder Name'], 'Shared'))
    for index, record in enumerate(records, start=5):
        ws.append(list(record))
        ws.row_dimensions[index].height = 46
        for cell in ws[index]:
            cell.alignment = Alignment(vertical='center', wrap_text=True)
            cell.font = Font(name='Aptos', size=11, color='0B1F33')
            if index % 2 == 0:
                cell.fill = PatternFill('solid', fgColor='EFF6F5')
        ws.cell(index, 1).font = Font(name='Aptos', bold=True, color='0B1F33')
        ws.cell(index, 2).font = Font(name='Aptos', color='008000', size=11)
    ws.auto_filter.ref = f'A4:C{ws.max_row}'
    meta = wb.create_sheet('Request Details')
    meta.sheet_view.showGridLines = False
    meta.append(['FIELD', 'VALUE'])
    info = [
        ('GTS ID', gts), ('Wordbee Order ID', order),
        ('Project Name', project), ('Requester', requester),
        ('Languages', ', '.join(languages)),
        ('Content Type', ' + '.join(x for x, yes in [('Marketing', marketing), ('Product', product)] if yes)),
    ]
    for item in info:
        meta.append(item)
    meta.column_dimensions['A'].width = 28
    meta.column_dimensions['B'].width = 75
    for cell in meta[1]:
        cell.fill = PatternFill('solid', fgColor='0B1F33')
        cell.font = Font(name='Aptos', color='FFFFFF', bold=True)
    for row in meta.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical='center', wrap_text=True)
        meta.row_dimensions[row[0].row].height = 30
    meta.freeze_panes = 'A2'
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def render_naming_converter():
    st.subheader('🏷️ Naming Converter')
    st.caption('Paste a Wordbee request, edit its name, and generate naming conventions using Tab 1 data.')
    raw = st.text_area('Paste Wordbee request',height=130,key='naming_raw',placeholder='Paste a Wordbee order row here...')
    parsed = parse_wordbee_text(raw)
    if st.session_state.get('naming_previous_raw') != raw:
        st.session_state['naming_previous_raw']=raw
        st.session_state['naming_order']=parsed['order']
        st.session_state['naming_project']=parsed['request']
        st.session_state['naming_requester']=parsed['requester']
        st.session_state.pop('naming_result',None)
    col1,col2=st.columns(2)
    with col1:
        order=st.text_input('Wordbee Order ID',key='naming_order')
        project=st.text_input('Project Name (editable)',key='naming_project')
    with col2:
        requester=st.text_input('Requester Name',key='naming_requester')
    # Preview the EDITABLE project field, not the originally pasted request.
    preview_name = project.strip()
    if raw.strip() or preview_name:
        st.markdown('<div style="border:1px solid rgba(127,127,127,.2);border-radius:16px;padding:16px;margin:8px 0"><div style="opacity:.7;font-size:.85rem">LIVE INPUT PREVIEW</div><div style="font-size:1.65rem;font-weight:700;overflow-wrap:anywhere">'+html.escape(preview_name or '—')+'</div></div>', unsafe_allow_html=True)
    gts=st.session_state.get('converter_gts') or st.session_state.get('shared_gts',get_default_project_code())
    langs=st.session_state.get('shared_languages',[])
    marketing=st.session_state.get('shared_marketing',False)
    product=st.session_state.get('shared_product',False)
    # Recover Tab 1 metadata from its uploader when the shared values are stale.
    tab1_file=st.session_state.get('converter_upload')
    if tab1_file is not None and (not langs or not (marketing or product)):
        try:
            tab1_marketing,tab1_product=process_all_sheets(io.BytesIO(tab1_file.getvalue()))
            langs=sorted(set(tab1_marketing.get('Language',pd.Series(dtype=str)).tolist()+tab1_product.get('Language',pd.Series(dtype=str)).tolist()))
            marketing=not tab1_marketing.empty
            product=not tab1_product.empty
        except Exception:
            pass
    st.markdown('#### Information from Tab 1')
    a,b,c=st.columns(3)
    a.text_input('GTS ID (from Tab 1)',value=gts,disabled=True)
    b.text_input('Languages (from Tab 1)',value=', '.join(langs) or 'None detected',disabled=True)
    c.text_input('Content Type (from Tab 1)',value=' + '.join([x for x,yes in [('Marketing',marketing),('Product',product)] if yes]) or 'None detected',disabled=True)
    if not langs or not (marketing or product):
        st.info('Upload a workbook in Tab 1 to detect languages and content type before converting.')
    valid_gts=bool(re.fullmatch(r'GTS\d+',gts.strip(),flags=re.I))
    ready=bool(re.fullmatch(r'\d+',order.strip()) and project.strip() and len(requester.strip().split())>=2 and valid_gts and langs and (marketing or product))
    if st.button('Convert',type='primary',disabled=not ready,key='naming_convert'):
        st.session_state['naming_result']=build_naming_results(order.strip(),project.strip(),requester.strip(),gts.strip(),langs,marketing,product)
        st.session_state['naming_result_signature']=(order,project,requester,gts,tuple(langs),marketing,product)
    signature=(order,project,requester,gts,tuple(langs),marketing,product)
    if st.session_state.get('naming_result_signature') != signature:
        st.session_state.pop('naming_result',None)
    result=st.session_state.get('naming_result')
    if result:
        st.success('Names generated successfully.')
        st.markdown('#### Wordbee Name')
        st.code(result['Wordbee Name'])
        st.markdown('#### AEM Names')
        for lang,value in result['AEM Names']:
            st.caption(lang)
            st.code(value)
        st.markdown('#### Shared Folder Name')
        st.code(result['Shared Folder Name'])
        excel_bytes = make_naming_excel(result, gts.strip(), order.strip(), project.strip(), requester.strip(), langs, marketing, product)
        st.download_button(
            '📥 Download Naming Convention Excel',
            data=excel_bytes,
            file_name=f'{gts.strip()} - Naming Convention.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            key='naming_excel_download',
        )
    if st.button('🔁 Reset',key='naming_reset'):
        reset_application()


def main():
    tab_converter,tab_aem,tab_naming=st.tabs(['🔗 URL Converter','🌐 AEM URLs','🏷️ Naming Converter'])
    with tab_converter:
        render_url_converter()
    with tab_aem:
        render_aem_urls()
    with tab_naming:
        render_naming_converter()


if __name__=='__main__':
    main()

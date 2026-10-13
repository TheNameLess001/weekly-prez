"""Session-owned files and state; never shared via a global Streamlit cache."""
from __future__ import annotations
import hashlib
import io
import json
import re
import tempfile
from pathlib import Path
from PIL import Image
from engine import FIELDS,inspect_file,ingest,dumps
from deck import DEFAULT,export_batch

DEFINITION_KEYS=('ca_definition','commission_definition','commission_base_definition','promo_definition','benefit_definition','line_gmv_definition','tax_basis')
THEME_COLORS=('primary','deep','background','white','text','good','bad','muted')
MAX_UPLOAD=100*1024*1024
MAX_SESSION=400*1024*1024

class Workspace:
    def __init__(self):
        self._temporary=tempfile.TemporaryDirectory(prefix='yassir_session_')
        self.root=Path(self._temporary.name)
        self.sources=[];self.settings={};self.theme=json.loads(json.dumps(DEFAULT));self.dataset=ingest([])
        self.revision=0;self.results=[];self.last_request=None;self.pending_config={}
    def close(self): self._temporary.cleanup()
    def reset(self):
        self.sources=[];self.settings={};self.theme=json.loads(json.dumps(DEFAULT));self.results=[];self.last_request=None;self.pending_config={};self.rebuild()
        # Existing files are only from this session, never another user's root.
        import shutil
        for child in self.root.iterdir():
            if child.is_dir():shutil.rmtree(child)
            else:child.unlink()
    def rebuild(self):
        self.dataset=ingest(self.sources,self.settings);self.settings=self.dataset.settings;self.revision+=1;self.results=[]
    def refresh_imports(self):
        """Upgrade pre-adapter sessions and saved mappings without discarding uploads."""
        for source in self.sources:
            for auto in inspect_file(source['path']):
                old=next((x for x in source['sheets'] if x['sheet']==auto['sheet']),None)
                if old is None:continue
                if auto.get('profile')=='admin_earnings':
                    old['mapping']={**auto['mapping'],**old.get('mapping',{})}
                    old['profile']='admin_earnings'
                    if old.get('grain')=='registry':old['grain']='orders'
                elif auto['grain']=='orders' and old.get('grain')=='registry' and not old.get('mapping',{}).get('date'):
                    old['mapping']={**auto['mapping'],**old.get('mapping',{})};old['grain']='orders'
        self.rebuild()

    def add_uploads(self,files):
        incoming=[];errors=[]
        total=sum(Path(s['path']).stat().st_size for s in self.sources if Path(s['path']).is_file())
        for name,raw in files:
            clean=Path(str(name).replace('\\','/')).name
            if Path(clean).suffix.lower() not in ('.csv','.xlsx'):
                errors.append(f'{clean}: format accepté .csv ou .xlsx');continue
            if len(raw)>MAX_UPLOAD or total+len(raw)>MAX_SESSION:
                errors.append(f'{clean}: limite 100 Mo par fichier / 400 Mo par session');continue
            digest=hashlib.sha256(raw).hexdigest();ident=hashlib.sha256((clean+digest).encode()).hexdigest()[:20]
            if any(s['id']==ident for s in self.sources):
                self.refresh_imports();continue
            folder=self.root/'imports'/digest[:16];folder.mkdir(parents=True,exist_ok=True);path=folder/clean;path.write_bytes(raw)
            try:sheets=inspect_file(path)
            except Exception as e:
                errors.append(f'{clean}: lecture impossible ({type(e).__name__})');path.unlink(missing_ok=True);continue
            saved=next((s for s in self.pending_config.get('source_configs',[]) if s.get('name')==clean),None)
            if saved:
                for sh in sheets:
                    old=next((x for x in saved.get('sheets',[]) if x.get('sheet')==sh['sheet']),None)
                    if old:
                        valid={k:v for k,v in old.get('mapping',{}).items() if k in FIELDS and v in sh['columns']}
                        sh['mapping'].update(valid);sh['grain']='orders' if sh.get('profile')=='admin_earnings' and old.get('grain')=='registry' else old.get('grain',sh['grain'])
                        if saved.get('digest')==digest:
                            for key in ('complete','coverage_start','coverage_end','all_partners','year','date_dayfirst','enabled'):sh[key]=old.get(key,sh.get(key))
            incoming.append({'id':ident,'name':clean,'digest':digest,'path':str(path),'sheets':sheets,'demo':any(sh.get('demo') for sh in sheets)})
            total+=len(raw)
        if incoming:
            if any(not s['demo'] for s in incoming) and any(s['demo'] for s in self.sources):
                self.sources=[s for s in self.sources if not s['demo']]
                # Do not apply fictitious financial definitions to real uploads.
                self.settings={k:v for k,v in self.settings.items() if k not in DEFINITION_KEYS or not str(v).startswith('DÉMO')}
            self.sources.extend(incoming);self.rebuild()
        return errors,len(incoming)
    def load_demo(self):
        from demo_data import generate,DEFINITIONS
        self.reset();folder=self.root/'demo';generate(folder)
        files=[(p.name,p.read_bytes()) for p in folder.glob('*.csv')]
        errors,n=self.add_uploads(files);self.settings=dict(DEFINITIONS)
        for source in self.sources:
            for sh in source['sheets']:
                if sh['grain']!='registry':sh.update(complete=True,coverage_start='2026-09-21',coverage_end='2026-10-04',all_partners=True)
        self.rebuild()
    def remove_source(self,ident):
        self.sources=[s for s in self.sources if s['id']!=ident];self.rebuild()
    def config_bytes(self,selection=None,period=None):
        source_configs=[{k:s[k] for k in ('name','digest','sheets')} for s in self.sources]
        for s in source_configs:
            s['sheets']=[{k:v for k,v in sh.items() if k not in ('sample','rows','demo')} for sh in s['sheets']]
        theme={k:self.theme[k] for k in (*THEME_COLORS,'font','currency')}
        return dumps({'schema_version':1,'source_configs':source_configs,'settings':self.settings,'theme':theme,
                      'selection':selection or [],'period':period or {}}).encode('utf-8')
    def apply_config(self,raw):
        if len(raw)>1024*1024:raise ValueError('Configuration limitée à 1 Mo')
        obj=json.loads(raw)
        if obj.get('schema_version')!=1:raise ValueError('Version de configuration inconnue')
        settings=obj.get('settings',{});sanitized={k:str(settings[k])[:2000] for k in DEFINITION_KEYS if k in settings}
        ids=settings.get('verified_ids',{})
        if not isinstance(ids,dict):raise ValueError('Correspondances ID invalides')
        sanitized['verified_ids']={str(k):str(v) for k,v in ids.items()}
        statuses=settings.get('statuses')
        if statuses is not None:
            if not isinstance(statuses,dict) or any(not isinstance(statuses.get(k,[]),list) or not all(isinstance(v,str) for v in statuses.get(k,[])) for k in ('delivered','cancelled','other')):raise ValueError('Statuts invalides')
            sanitized['statuses']={k:statuses.get(k,[]) for k in ('delivered','cancelled','other')}
        theme=self.clean_theme(obj.get('theme',{}))
        configs=obj.get('source_configs',[])
        if not isinstance(configs,list):raise ValueError('Mappings invalides')
        for config in configs:
            if not isinstance(config,dict):raise ValueError('Mapping invalide')
            for sh in config.get('sheets',[]):
                if sh.get('grain') not in ('orders','items','daily','weekly','period','registry','complement'):raise ValueError('Grain invalide')
                if not isinstance(sh.get('mapping',{}),dict):raise ValueError('Mapping invalide')
        self.settings=sanitized;self.theme.update(theme);self.pending_config={**obj,'source_configs':configs}
        for source in self.sources:
            saved=next((s for s in configs if s.get('name')==source['name']),None)
            if not saved:continue
            for sh in source['sheets']:
                old=next((x for x in saved.get('sheets',[]) if x.get('sheet')==sh['sheet']),None)
                if not old:continue
                sh['mapping'].update({k:v for k,v in old.get('mapping',{}).items() if k in FIELDS and v in sh['columns']})
                sh['grain']=old.get('grain',sh['grain'])
                if saved.get('digest')==source['digest']:
                    for k in ('complete','coverage_start','coverage_end','all_partners','year','date_dayfirst','enabled'):sh[k]=old.get(k,sh.get(k))
                else:sh.update(complete=False,coverage_start='',coverage_end='')
        self.rebuild();return obj
    @staticmethod
    def clean_theme(value):
        out={}
        for k in THEME_COLORS:
            if k in value:
                if not re.fullmatch(r'#[0-9a-fA-F]{6}',str(value[k])):raise ValueError('Couleur invalide : '+k)
                out[k]=value[k]
        for k in ('font','currency'):
            if k in value:
                if not str(value[k]).strip():raise ValueError('Valeur obligatoire : '+k)
                out[k]=str(value[k]).strip()[:50]
        # Server file paths are never imported through a downloaded config.
        return out
    def save_logo(self,name,raw,rid=None):
        if len(raw)>5*1024*1024:raise ValueError('Logo limité à 5 Mo')
        if str(name).lower().endswith('.svg'):
            if re.search(rb'@import|url\(\s*[\'"]?(?!#)|<!DOCTYPE',raw,re.I):raise ValueError('Ressource SVG externe refusée')
            from xml.etree import ElementTree as ET
            from io import BytesIO
            tree=ET.fromstring(raw)
            for node in tree.iter():
                if node.tag.split('}')[-1].lower() in ('script','foreignobject'):raise ValueError('SVG actif refusé')
                for k,v in node.attrib.items():
                    if k.split('}')[-1].lower().startswith('on') or (k.split('}')[-1]=='href' and not v.startswith('#')):raise ValueError('Ressource SVG externe refusée')
            import cairosvg
            raw=cairosvg.svg2png(bytestring=raw,output_width=1600)
        with Image.open(io.BytesIO(raw)) as image:
            if image.width*image.height>24_000_000:raise ValueError('Logo trop grand')
            image=image.convert('RGBA');image.thumbnail((1600,1600))
            folder=self.root/'logos';folder.mkdir(exist_ok=True);path=folder/(hashlib.sha256(raw).hexdigest()[:20]+'.png');image.save(path)
        if rid:self.theme.setdefault('restaurant_logos',{})[rid]=str(path)
        else:self.theme['yassir_logo']=str(path)
        self.results=[];return path
    def generate(self,ids,current,previous,pdf=False,progress=None):
        result=export_batch(self.dataset,ids,current,previous,self.root/'exports',self.theme,pdf,progress)
        request={'ids':list(ids),'current':current,'previous':previous,'pdf':pdf}
        result['request']=request;self.results.append(result);self.last_request=request
        return result

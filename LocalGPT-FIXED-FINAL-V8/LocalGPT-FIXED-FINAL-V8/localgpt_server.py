import os, re, json, socket, threading, urllib.request, urllib.error, uuid, base64, mimetypes, time, subprocess
from pathlib import Path
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent
DATA = ROOT / 'data'
UPLOADS = DATA / 'uploads'; TEXTS = DATA / 'texts'; META = DATA / 'meta'; CHATS = DATA / 'chats'
for p in (UPLOADS, TEXTS, META, CHATS): p.mkdir(parents=True, exist_ok=True)
HOST='127.0.0.1'; DEFAULT_PORT=int(os.environ.get('LOCALGPT_PORT','3000'))
OLLAMA_URL=os.environ.get('OLLAMA_URL','http://127.0.0.1:11434'); MODEL=os.environ.get('LOCALGPT_MODEL','llama3.2:latest')
MAX_CONTEXT=24000
IMAGE_EXT={'.png','.jpg','.jpeg','.webp','.bmp','.gif'}
SUPPORTED={'.pdf','.docx','.pptx','.xlsx','.txt','.md','.csv','.json'}|IMAGE_EXT

def clean_name(name): return re.sub(r'[^A-Za-z0-9._ -]','_',os.path.basename(name or 'file'))[:120] or 'file'
def meta_path(fid): return META/(fid+'.json')
def text_path(fid): return TEXTS/(fid+'.txt')
def chat_path(cid): return CHATS/(cid+'.json')
def save_meta(m): meta_path(m['id']).write_text(json.dumps(m,ensure_ascii=False,indent=2),encoding='utf-8')
def load_meta(fid):
    try:
        p=meta_path(fid); return json.loads(p.read_text(encoding='utf-8')) if p.exists() else None
    except: return None
def all_meta():
    out=[]
    for p in META.glob('*.json'):
        try: out.append(json.loads(p.read_text(encoding='utf-8')))
        except: pass
    return sorted(out,key=lambda x:x.get('created_at',0))
def read_doc(fid):
    m=load_meta(fid)
    if not m:return None,''
    try:return m,text_path(fid).read_text(encoding='utf-8',errors='ignore')
    except:return m,''
def save_chat(cid,obj): chat_path(cid).write_text(json.dumps(obj,ensure_ascii=False,indent=2),encoding='utf-8')
def all_chats():
    out=[]
    for p in CHATS.glob('*.json'):
        try: out.append(json.loads(p.read_text(encoding='utf-8')))
        except: pass
    return sorted(out,key=lambda x:x.get('updated_at',x.get('created_at',0)),reverse=True)

def load_chat(cid):
    try:
        p=chat_path(cid)
        if p.exists(): return json.loads(p.read_text(encoding='utf-8'))
    except: pass
    return {'id':cid,'title':'New chat','file_ids':[],'messages':[],'created_at':time.time(),'updated_at':time.time()}

def extract_pdf(path):
    try:
        import fitz
        d=fitz.open(path); pages=[]
        for i,page in enumerate(d,1):
            txt=page.get_text('text').strip()
            if txt: pages.append(f'[PAGE {i}]\n{txt}')
        d.close(); return '\n\n'.join(pages).strip()
    except:return ''
def extract_docx(path):
    try:
        with zipfile.ZipFile(path) as z: root=ET.fromstring(z.read('word/document.xml'))
        ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        return '\n'.join(''.join(t.text or '' for t in p.findall('.//w:t',ns)) for p in root.findall('.//w:p',ns)).strip()
    except:return ''
def extract_pptx(path):
    try:
        with zipfile.ZipFile(path) as z:
            ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}; out=[]
            for n in sorted(x for x in z.namelist() if x.startswith('ppt/slides/slide') and x.endswith('.xml')):
                r=ET.fromstring(z.read(n)); out.append(' '.join((t.text or '') for t in r.findall('.//a:t',ns)))
            return '\n'.join(out).strip()
    except:return ''
def extract_xlsx(path):
    try:
        with zipfile.ZipFile(path) as z:
            ns={'a':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}; shared=[]
            if 'xl/sharedStrings.xml' in z.namelist():
                r=ET.fromstring(z.read('xl/sharedStrings.xml')); shared=[''.join(t.text or '' for t in si.findall('.//a:t',ns)) for si in r.findall('.//a:si',ns)]
            out=[]
            for n in sorted(x for x in z.namelist() if x.startswith('xl/worksheets/sheet') and x.endswith('.xml')):
                r=ET.fromstring(z.read(n)); rows=[]
                for row in r.findall('.//a:row',ns):
                    vals=[]
                    for c in row.findall('a:c',ns):
                        v=c.find('a:v',ns)
                        if v is None: continue
                        val=v.text or ''
                        if c.attrib.get('t')=='s':
                            try: val=shared[int(val)]
                            except: pass
                        vals.append(val)
                    if vals: rows.append(' | '.join(vals))
                out.append('\n'.join(rows))
            return '\n'.join(out).strip()
    except:return ''
def extract_image(path):
    try:
        from PIL import Image
        img=Image.open(path)
        # First try Tesseract if it is already installed.
        try:
            import pytesseract
            text=pytesseract.image_to_string(img).strip()
            if text:return '[IMAGE OCR TEXT]\n'+text
        except Exception:
            pass
        # Windows 10/11 has a built-in local OCR engine. This avoids downloading
        # another OCR package and keeps the document local.
        try:
            script=ROOT/'windows_ocr.ps1'
            if script.exists():
                r=subprocess.run(['powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-File',str(script),str(path)],capture_output=True,text=True,encoding='utf-8',errors='ignore',timeout=45)
                text=(r.stdout or '').strip()
                if r.returncode==0 and text:
                    return '[IMAGE OCR TEXT]\n'+text
        except Exception:
            pass
        return f'[IMAGE FILE]\nImage size: {img.width}x{img.height}. No readable text was detected by local OCR.'
    except Exception:
        return '[IMAGE FILE]\nThe image was uploaded, but local image reading is unavailable.'
def extract_text(path):
    e=path.suffix.lower()
    if e=='.pdf':return extract_pdf(path)
    if e=='.docx':return extract_docx(path)
    if e=='.pptx':return extract_pptx(path)
    if e=='.xlsx':return extract_xlsx(path)
    if e in IMAGE_EXT:return extract_image(path)
    if e in {'.txt','.md','.csv','.json'}:
        try:return path.read_text(encoding='utf-8',errors='ignore').strip()
        except:return ''
    return ''

def normalize(s):
    s=re.sub(r'<think>.*?</think>','',s or '',flags=re.S|re.I)
    return re.sub(r'\n{3,}','\n\n',s).strip()
def ollama_models():
    try:
        with urllib.request.urlopen(OLLAMA_URL+'/api/tags',timeout=4) as r:return json.loads(r.read()).get('models',[])
    except:return []
def vision_model():
    for x in ollama_models():
        n=x.get('name','').lower()
        if 'vision' in n:return x.get('name')
    return None
def ollama_generate(prompt, images=None, model=None):
    payload={'model':model or MODEL,'prompt':prompt,'stream':False,'keep_alive':'10m','options':{'temperature':0.15,'num_ctx':6144,'num_predict':384}}
    if images: payload['images']=images
    req=urllib.request.Request(OLLAMA_URL+'/api/generate',data=json.dumps(payload).encode(),headers={'Content-Type':'application/json'})
    with urllib.request.urlopen(req,timeout=180) as r:return normalize(json.loads(r.read()).get('response',''))

def direct_fact(q,context):
    ql=q.lower()
    if re.search(r'\b(cgpa|c\.g\.p\.a|gpa)\b',ql):
        for p in [r'(?:cgpa|c\.g\.p\.a|gpa)\s*[:\-]?\s*(\d+(?:\.\d+)?)',r'(\d+(?:\.\d+)?)\s*(?:cgpa|gpa)']:
            m=re.search(p,context,re.I)
            if m:return f'Your CGPA is {m.group(1)}.'
    if re.search(r'\b(email|e-mail)\b',ql):
        m=re.search(r'\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b',context,re.I)
        if m:return f'Your email is {m.group(0)}.'
    if re.search(r'\b(phone|mobile|contact number)\b',ql):
        n=re.findall(r'(?<!\d)(?:\+91[\s-]?)?[6-9]\d{9}(?!\d)',context)
        if n:return f'Your phone number is {n[0]}.'
    return None

def combined_context(file_ids):
    blocks=[]; imgs=[]
    for fid in file_ids:
        m,t=read_doc(fid)
        if not m:continue
        blocks.append(f"\n===== FILE: {m['display_name']} =====\n{t[:MAX_CONTEXT//max(1,len(file_ids))]}")
        if m.get('is_image'):
            p=UPLOADS/m['stored_name']
            try: imgs.append((m,base64.b64encode(p.read_bytes()).decode()))
            except: pass
    return '\n'.join(blocks),imgs

def answer(file_ids,q):
    context,imgs=combined_context(file_ids)
    if not context:
        prompt=f'''You are LocalGPT, a private local AI assistant. Answer the user's question naturally and directly. No files are attached to this chat, so do not invent file-specific information. Never mention backend, RAG, retrieval, pipeline, or internal processing.

USER QUESTION:
{q}'''
        return ollama_generate(prompt) or 'I could not generate an answer.'
    if len(file_ids)==1:
        fact=direct_fact(q,context)
        if fact:return fact
    prompt=f'''You are LocalGPT, a private local AI assistant. Answer the user's question using ONLY the files attached to the CURRENT CHAT.
Never use, mention, or mix files from other chats. If multiple files are attached here, compare/use them only when the question requires it.
For resume/job questions, use the candidate's actual education, skills, projects and experience from the attached resume; give practical role suggestions without inventing qualifications.
For summaries, summarize only the attached current-chat files.
Use the attached files whenever they are relevant to the user's question. If the files do not contain the requested information, answer the general question using your normal knowledge instead of refusing. Clearly say when the answer is general knowledge rather than information found in the attached files. Never invent facts about the attached files. For page-specific questions, use the [PAGE N] markers when present. Answer directly and concisely. Never mention backend, RAG, retrieval, pipeline, sub-queries, or internal processing.

USER QUESTION:
{q}

CURRENT CHAT FILES:
{context[:MAX_CONTEXT]}
'''
    vm=vision_model()
    if any(m.get('is_image') for m,_ in imgs) and vm:
        # Use vision only when the user has a local vision model installed.
        return ollama_generate(prompt,[b for _,b in imgs],model=vm) or "I couldn't find that in the files attached to this chat."
    return ollama_generate(prompt) or "I couldn't find that in the files attached to this chat."

def summarize(file_ids):
    context,_=combined_context(file_ids)
    if not context:return 'No readable content is available in this chat.'
    prompt=f'''Summarize ONLY the files attached to the CURRENT CHAT. Do not use information from any other chat. If there are multiple files, give a clear summary of each file and then a short combined overview. Do not invent details.\n\n{context[:MAX_CONTEXT]}'''
    return ollama_generate(prompt) or 'No summary could be generated.'

class Handler(BaseHTTPRequestHandler):
    server_version='LocalGPT/3.0'
    def log_message(self,fmt,*args): print('[HTTP]',fmt%args)
    def send_data(self,code,data,ctype='application/json; charset=utf-8'):
        raw=data.encode() if isinstance(data,str) else data; self.send_response(code); self.send_header('Content-Type',ctype); self.send_header('Content-Length',str(len(raw))); self.send_header('Cache-Control','no-store'); self.end_headers(); self.wfile.write(raw)
    def j(self,code,obj):self.send_data(code,json.dumps(obj,ensure_ascii=False))
    def do_OPTIONS(self):self.send_response(204); self.send_header('Access-Control-Allow-Origin','*'); self.send_header('Access-Control-Allow-Headers','Content-Type'); self.send_header('Access-Control-Allow-Methods','GET,POST,DELETE,OPTIONS'); self.end_headers()
    def read_json(self):
        n=int(self.headers.get('Content-Length','0')); return json.loads((self.rfile.read(n) if n else b'{}').decode('utf-8','ignore') or '{}')
    def do_GET(self):
        path=self.path.split('?',1)[0]
        if path=='/api/status':return self.j(200,{'ok':True,'ollama':bool(ollama_models()),'model':MODEL,'vision_model':vision_model(),'files':all_meta()})
        if path=='/api/chats':return self.j(200,{'chats':all_chats()})
        if path.startswith('/api/chat/'):return self.j(200,load_chat(path.rsplit('/',1)[-1]))
        if path=='/':return self.send_data(200,(ROOT/'index.html').read_bytes(),'text/html; charset=utf-8')
        return self.send_data(404,'Not found','text/plain')
    def do_POST(self):
        if self.path=='/api/upload':return self.upload()
        if self.path=='/api/chat':return self.chat()
        if self.path=='/api/new-chat':
            cid=uuid.uuid4().hex; c={'id':cid,'title':'New chat','file_ids':[],'messages':[],'created_at':time.time(),'updated_at':time.time()}; save_chat(cid,c); return self.j(200,c)
        if self.path=='/api/update-chat':return self.update_chat()
        if self.path=='/api/summarize':return self.summary()
        if self.path=='/api/clear-chat':return self.clear_chat()
        return self.j(404,{'error':'Not found'})
    def do_DELETE(self):
        path=self.path.split('?',1)[0]
        if path.startswith('/api/files/'):
            fid=path.rsplit('/',1)[-1]; m=load_meta(fid)
            if not m:return self.j(404,{'error':'File not found'})
            for p in ((UPLOADS/m['stored_name']),text_path(fid),meta_path(fid)):
                try:p.unlink(missing_ok=True)
                except:pass
            return self.j(200,{'ok':True})
        if path.startswith('/api/chats/'):
            cid=path.rsplit('/',1)[-1]
            try:chat_path(cid).unlink(missing_ok=True)
            except:pass
            return self.j(200,{'ok':True})
        return self.j(404,{'error':'Not found'})
    def upload(self):
        ctype=self.headers.get('Content-Type','')
        if not ctype.startswith('multipart/form-data'):return self.j(400,{'error':'Invalid upload request'})
        try:
            from email.parser import BytesParser
            from email.policy import default
            body=self.rfile.read(int(self.headers.get('Content-Length','0')))
            msg=BytesParser(policy=default).parsebytes((f'Content-Type: {ctype}\r\nMIME-Version: 1.0\r\n\r\n').encode()+body)
            saved=[]
            for item in list(msg.iter_parts()):
                name=item.get_filename() if item.get_param('name',header='content-disposition')=='files' else None
                name=clean_name(name or '')
                if not name or Path(name).suffix.lower() not in SUPPORTED:continue
                data=item.get_payload(decode=True) or b''
                if not data:continue
                fid=uuid.uuid4().hex; ext=Path(name).suffix.lower(); stored=fid+ext; p=UPLOADS/stored; p.write_bytes(data)
                text=extract_text(p)
                if not text and ext not in IMAGE_EXT: p.unlink(missing_ok=True); continue
                text_path(fid).write_text(text,encoding='utf-8')
                m={'id':fid,'display_name':name,'stored_name':stored,'extension':ext,'created_at':time.time(),'characters':len(text),'is_image':ext in IMAGE_EXT}; save_meta(m); saved.append(m)
            if not saved:return self.j(400,{'error':'No readable files. Supported: PDF, DOCX, PPTX, XLSX, TXT, MD, CSV, JSON, PNG, JPG, JPEG, WEBP, BMP, GIF.'})
            return self.j(200,{'ok':True,'files':saved})
        except Exception as e:return self.j(500,{'error':'Upload failed: '+str(e)})
    def chat(self):
        try:
            d=self.read_json(); cid=str(d.get('chat_id','')).strip(); q=str(d.get('query','')).strip()
            if not cid or not q:return self.j(400,{'error':'Create a chat and enter a question.'})
            c=load_chat(cid)
            # The server is authoritative: a chat may use ONLY the files already
            # attached to that chat. Never trust file_ids supplied by the browser.
            fids=[str(x) for x in c.get('file_ids',[]) if load_meta(str(x))]
            # Persist the user's message BEFORE running the local model.
            # This is important: switching chats while Ollama is thinking must
            # never make the just-sent message disappear or leak into another chat.
            c['file_ids']=fids
            if c.get('title')=='New chat':
                c['title']=q[:42] if q else 'New chat'
            messages=c.get('messages',[])
            messages.append({'role':'user','content':q})
            c['messages']=messages
            c['updated_at']=time.time()
            save_chat(cid,c)

            try:
                ans=answer(fids,q)
            except Exception as model_error:
                ans='Local AI error: '+str(model_error)

            # Append the answer to THIS chat only.
            c=load_chat(cid)
            c['file_ids']=fids
            c['messages']=c.get('messages',[]) + [{'role':'assistant','content':ans}]
            c['updated_at']=time.time()
            save_chat(cid,c)
            return self.j(200,{'ok':True,'response':ans})
        except Exception as e:return self.j(500,{'error':str(e)})
    def update_chat(self):
        try:
            d=self.read_json(); cid=str(d.get('chat_id','')).strip(); c=load_chat(cid)
            if 'title' in d:c['title']=str(d.get('title') or 'New chat')[:80]
            if 'file_ids' in d:c['file_ids']=[str(x) for x in d.get('file_ids',[]) if str(x) and load_meta(str(x))]
            c['updated_at']=time.time(); save_chat(cid,c); return self.j(200,c)
        except Exception as e:return self.j(500,{'error':str(e)})

    def summary(self):
        try:
            d=self.read_json(); cid=str(d.get('chat_id','')).strip(); c=load_chat(cid); fids=c.get('file_ids',[])
            if not fids:return self.j(400,{'error':'Attach files to this chat first.'})
            ans=summarize(fids); c['messages']=c.get('messages',[])+[{'role':'assistant','content':ans,'type':'summary'}]; c['updated_at']=time.time(); save_chat(cid,c); return self.j(200,{'ok':True,'response':ans})
        except Exception as e:return self.j(500,{'error':str(e)})
    def clear_chat(self):
        try:
            d=self.read_json(); cid=str(d.get('chat_id','')).strip(); c=load_chat(cid); c['messages']=[]; c['updated_at']=time.time(); save_chat(cid,c); return self.j(200,{'ok':True})
        except Exception as e:return self.j(500,{'error':str(e)})

def find_port():
    for p in range(DEFAULT_PORT,DEFAULT_PORT+10):
        s=socket.socket()
        try:s.bind((HOST,p));s.close();return p
        except OSError:s.close()
    raise RuntimeError('Ports 3000-3009 are busy.')
if __name__=='__main__':
    port=find_port(); print(f'LocalGPT ready: http://localhost:{port}'); print(f'AI model: {MODEL}'); print(f'Vision model: {vision_model() or "OCR only"}')
    httpd=ThreadingHTTPServer((HOST,port),Handler); httpd.daemon_threads=True
    try: threading.Timer(.8,lambda:__import__('webbrowser').open(f'http://localhost:{port}')).start()
    except:pass
    httpd.serve_forever()

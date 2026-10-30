#!/usr/bin/env python3
"""쩡마술사 링크 모음 - 로컬 서버

  python server.py              # http://127.0.0.1:8765 에서 열림
  python server.py --model bge-m3 --port 8765 --data links.json

하는 일
  - index.html 을 서빙하고, 링크 목록을 links.json 에 저장
  - 링크 페이지의 제목/설명/대표 이미지(og 태그)를 가져와 썸네일로 축소
  - Ollama 임베딩(기본 bge-m3)으로 관련 자료 찾기용 벡터 생성

표준 라이브러리만으로 동작한다. Pillow 가 있으면 썸네일을 작게 줄여
JSON 안에 넣으므로(백업을 아티팩트로 옮겨도 그대로 보임) 설치를 권장한다.
  pip install pillow
  ollama pull bge-m3
"""
import argparse
import base64
import html
import io
import json
import os
import re
import struct
import threading
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import zlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

try:
    from PIL import Image
except ImportError:  # 썸네일은 원본 URL 로 대체
    Image = None

HERE = os.path.dirname(os.path.abspath(__file__))
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")
THUMB_W, THUMB_H = 240, 160

ARGS = None
STORE_LOCK = threading.Lock()


# ---------------------------------------------------------------- 저장소
def load_items():
    try:
        with open(ARGS.data, encoding="utf-8") as f:
            data = json.load(f)
        return data.get("items", data) if isinstance(data, dict) else data
    except FileNotFoundError:
        return []


def save_items(items):
    tmp = ARGS.data + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"v": 1, "items": items}, f, ensure_ascii=False)
    os.replace(tmp, ARGS.data)


def store_put(item_id, item, merge=False):
    with STORE_LOCK:
        items = load_items()
        for i, it in enumerate(items):
            if it.get("id") == item_id:
                items[i] = {**it, **item} if merge else {**item, "id": item_id}
                break
        else:
            if merge:
                return False
            items.append({**item, "id": item_id})
        save_items(items)
        return True


def store_delete(item_id):
    with STORE_LOCK:
        items = [it for it in load_items() if it.get("id") != item_id]
        save_items(items)


# ---------------------------------------------------------------- HTTP 가져오기
def http_get(url, limit=2_000_000, timeout=10, accept="text/html,application/xhtml+xml,*/*;q=0.8"):
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Accept": accept,
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
        "Accept-Encoding": "gzip, deflate",
    })
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = r.read(limit)
        enc = (r.headers.get("Content-Encoding") or "").lower()
        ctype = r.headers.get("Content-Type") or ""
        final = r.geturl()
    if enc == "gzip":
        data = zlib.decompressobj(16 + zlib.MAX_WBITS).decompress(data)  # 잘린 스트림도 허용
    elif enc == "deflate":
        try:
            data = zlib.decompressobj().decompress(data)
        except zlib.error:
            data = zlib.decompressobj(-zlib.MAX_WBITS).decompress(data)
    return data, ctype, final


def decode_html(data, ctype):
    m = re.search(r"charset=([\w-]+)", ctype, re.I)
    cs = m.group(1) if m else None
    if not cs:
        m = re.search(rb"<meta[^>]+charset=[\"']?([\w-]+)", data[:4096], re.I)
        cs = m.group(1).decode("ascii", "ignore") if m else "utf-8"
    cs = cs.lower()
    if cs in ("euc-kr", "ks_c_5601-1987", "ksc5601"):
        cs = "cp949"
    try:
        return data.decode(cs, errors="replace")
    except LookupError:
        return data.decode("utf-8", errors="replace")


META_RE = re.compile(r"<meta\s[^>]*>", re.I)
ATTR_RE = re.compile(r"([\w:-]+)\s*=\s*(?:\"([^\"]*)\"|'([^']*)'|([^\s>]+))")


def parse_meta(doc):
    meta = {}
    head = doc[:400_000]
    for tag in META_RE.findall(head):
        attrs = {m.group(1).lower(): (m.group(2) or m.group(3) or m.group(4) or "")
                 for m in ATTR_RE.finditer(tag)}
        key = (attrs.get("property") or attrs.get("name") or attrs.get("itemprop") or "").lower()
        if key and "content" in attrs and key not in meta:
            meta[key] = html.unescape(attrs["content"]).strip()
    m = re.search(r"<title[^>]*>(.*?)</title>", head, re.I | re.S)
    if m:
        meta["<title>"] = html.unescape(re.sub(r"\s+", " ", m.group(1))).strip()
    return meta


def pick(meta, *keys):
    for k in keys:
        if meta.get(k):
            return meta[k]
    return ""


def clean(s, n):
    s = re.sub(r"\s+", " ", s or "").strip()
    return s[:n] + ("…" if len(s) > n else "")


def make_thumb(img_url):
    """대표 이미지를 내려받아 작은 JPEG data URI 로. Pillow 가 없으면 URL 그대로."""
    if not img_url:
        return ""
    if Image is None:
        return img_url
    try:
        data, _, _ = http_get(img_url, limit=8_000_000, accept="image/*,*/*;q=0.5")
        im = Image.open(io.BytesIO(data))
        im.draft("RGB", (THUMB_W * 2, THUMB_H * 2))
        im = im.convert("RGB")
        # 가운데 기준으로 잘라 비율 맞추기
        w, h = im.size
        tr = THUMB_W / THUMB_H
        if w / h > tr:
            nw = int(h * tr)
            im = im.crop(((w - nw) // 2, 0, (w - nw) // 2 + nw, h))
        else:
            nh = int(w / tr)
            im = im.crop((0, (h - nh) // 2, w, (h - nh) // 2 + nh))
        im = im.resize((THUMB_W, THUMB_H), Image.LANCZOS)
        out = io.BytesIO()
        im.save(out, "JPEG", quality=72, optimize=True, progressive=True)
        return "data:image/jpeg;base64," + base64.b64encode(out.getvalue()).decode()
    except Exception:
        return ""


def youtube_id(url):
    p = urllib.parse.urlparse(url)
    host = p.hostname or ""
    if host.endswith("youtu.be"):
        return p.path.strip("/").split("/")[0]
    if "youtube.com" in host:
        q = urllib.parse.parse_qs(p.query)
        if q.get("v"):
            return q["v"][0]
        m = re.match(r"/(?:shorts|live|embed)/([\w-]+)", p.path)
        if m:
            return m.group(1)
    return ""


def fetch_meta(url):
    title = summary = site = img = ""
    target = url
    # 사이트별 보정
    m = re.match(r"https?://(?:m\.)?blog\.naver\.com/([^/?#]+)/(\d+)", url)
    if m:  # PC 블로그는 본문이 iframe 이라 og 태그가 빈약하다
        target = f"https://m.blog.naver.com/{m.group(1)}/{m.group(2)}"
    yid = youtube_id(url)
    if yid:
        try:
            data, _, _ = http_get("https://www.youtube.com/oembed?format=json&url=" +
                                  urllib.parse.quote(f"https://www.youtube.com/watch?v={yid}"),
                                  accept="application/json")
            j = json.loads(data)
            title, site = j.get("title", ""), "YouTube · " + j.get("author_name", "")
        except Exception:
            pass
        img = f"https://i.ytimg.com/vi/{yid}/hqdefault.jpg"

    if not title:
        try:
            data, ctype, final = http_get(target)
            if "html" in ctype.lower() or not ctype:
                meta = parse_meta(decode_html(data, ctype))
                title = pick(meta, "og:title", "twitter:title", "<title>")
                summary = pick(meta, "og:description", "twitter:description", "description")
                site = pick(meta, "og:site_name", "application-name")
                img = img or pick(meta, "og:image:secure_url", "og:image", "twitter:image", "twitter:image:src", "image")
                if img:
                    img = urllib.parse.urljoin(final, img)
            elif ctype.lower().startswith("image/"):
                img = final
        except urllib.error.HTTPError as e:
            if not re.search(r"reddit\.com", url):
                raise RuntimeError(f"HTTP {e.code}")
        except Exception as e:
            if not re.search(r"reddit\.com", url):
                raise RuntimeError(str(e)[:120])

    if not title and "reddit.com" in url:  # reddit 은 봇 차단이 잦아 oEmbed 로
        try:
            data, _, _ = http_get("https://www.reddit.com/oembed?url=" + urllib.parse.quote(url),
                                  accept="application/json")
            j = json.loads(data)
            title, site = j.get("title", ""), "Reddit"
        except Exception:
            pass

    return {
        "title": clean(title, 160),
        "summary": clean(summary, 240),
        "site": clean(site, 60),
        "thumb": make_thumb(img),
    }


# ---------------------------------------------------------------- Ollama
def ollama(path, body=None, timeout=60):
    url = ARGS.ollama.rstrip("/") + path
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read())


def ollama_status():
    try:
        tags = ollama("/api/tags", timeout=2)
    except Exception:
        return "off"
    names = {t.get("name", "") for t in tags.get("models", [])}
    want = ARGS.model
    if want in names or f"{want}:latest" in names or any(n.split(":")[0] == want for n in names):
        return "ready"
    return "missing"


def embed(text):
    """정규화 후 int8 양자화한 벡터를 base64 로. 코사인 유사도는 스케일과 무관하다."""
    try:
        vec = ollama("/api/embed", {"model": ARGS.model, "input": text})["embeddings"][0]
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
        vec = ollama("/api/embeddings", {"model": ARGS.model, "prompt": text})["embedding"]
    peak = max(abs(v) for v in vec) or 1.0
    q = [max(-127, min(127, round(v / peak * 127))) for v in vec]
    return base64.b64encode(struct.pack(f"{len(q)}b", *q)).decode()


# ---------------------------------------------------------------- HTTP 서버
class Handler(BaseHTTPRequestHandler):
    server_version = "LinkArchive/1"

    def log_message(self, fmt, *a):
        if ARGS.verbose:
            super().log_message(fmt, *a)

    def send_json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def read_json(self):
        n = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(n) or b"{}")

    def item_id(self):
        m = re.fullmatch(r"/api/items/([A-Za-z0-9_\-.~:@+]{1,200})", self.path)
        return m.group(1) if m else None

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            with open(os.path.join(HERE, "index.html"), "rb") as f:
                body = f.read()
            # 아티팩트용 파일이라 doctype 이 없다. 표준 모드로 렌더하도록 감싼다.
            body = (b"<!doctype html><html lang=\"ko\"><head><meta charset=\"utf-8\">"
                    b"<meta name=\"viewport\" content=\"width=device-width,initial-scale=1,viewport-fit=cover\">"
                    b"<style>body{margin:0}[hidden]{display:none!important}img{max-width:100%}</style></head><body>"
                    + body + b"</body></html>")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)
        elif self.path == "/api/ping":
            self.send_json({"ok": True, "ollama": ollama_status(), "model": ARGS.model,
                            "pillow": Image is not None})
        elif self.path == "/api/items":
            with STORE_LOCK:
                self.send_json({"items": load_items()})
        else:
            self.send_json({"error": "not found"}, 404)

    def do_PUT(self):
        iid = self.item_id()
        if not iid:
            return self.send_json({"error": "bad id"}, 400)
        store_put(iid, self.read_json())
        self.send_json({"ok": True})

    def do_PATCH(self):
        iid = self.item_id()
        if not iid:
            return self.send_json({"error": "bad id"}, 400)
        ok = store_put(iid, self.read_json(), merge=True)
        self.send_json({"ok": ok}, 200 if ok else 404)

    def do_DELETE(self):
        iid = self.item_id()
        if not iid:
            return self.send_json({"error": "bad id"}, 400)
        store_delete(iid)
        self.send_json({"ok": True})

    def do_POST(self):
        try:
            body = self.read_json()
            if self.path == "/api/enrich":
                url = body.get("url", "")
                if not re.match(r"https?://", url):
                    return self.send_json({"error": "http(s) 주소만 가져올 수 있습니다"}, 400)
                self.send_json(fetch_meta(url))
            elif self.path == "/api/embed":
                text = (body.get("text") or "").strip()
                if not text:
                    return self.send_json({"error": "빈 텍스트"}, 400)
                self.send_json({"emb": embed(text[:4000]), "model": ARGS.model})
            else:
                self.send_json({"error": "not found"}, 404)
        except Exception as e:
            self.send_json({"error": str(e)[:200]}, 502)


def main():
    global ARGS
    ap = argparse.ArgumentParser(description="쩡마술사 링크 모음 로컬 서버")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--data", default=os.path.join(HERE, "links.json"), help="링크 저장 파일")
    ap.add_argument("--model", default="bge-m3", help="Ollama 임베딩 모델 (한국어는 bge-m3 권장)")
    ap.add_argument("--ollama", default=os.environ.get("OLLAMA_HOST", "http://127.0.0.1:11434"))
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    ARGS = ap.parse_args()
    if not ARGS.ollama.startswith("http"):
        ARGS.ollama = "http://" + ARGS.ollama

    srv = ThreadingHTTPServer(("127.0.0.1", ARGS.port), Handler)
    url = f"http://127.0.0.1:{ARGS.port}/"
    print(f"링크 모음: {url}")
    print(f"  저장 파일 : {ARGS.data}")
    print(f"  썸네일    : {'Pillow 로 축소' if Image else '원본 URL (pip install pillow 권장)'}")
    st = ollama_status()
    print(f"  Ollama    : {ARGS.model} " + {"ready": "준비됨", "missing": f"모델 없음 → ollama pull {ARGS.model}",
                                             "off": "꺼져 있음 (글자 비교로 관련 자료를 찾습니다)"}[st])
    if not ARGS.no_browser:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()

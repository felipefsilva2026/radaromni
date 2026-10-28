"""Descobre o endereço real da matéria e a foto de destaque dela.

Os links do RSS do Google Notícias (news.google.com/rss/articles/...) não levam direto
ao site da matéria: o Google faz o redirecionamento via JavaScript. Aqui o link é
convertido no endereço real pela mesma chamada que a página do Google faz, e depois
lemos a foto de destaque que o próprio site publica para compartilhamento
(meta tag og:image / twitter:image — a imagem que aparece ao mandar o link no WhatsApp).

Tudo é "melhor esforço": se algo falhar, a notícia fica com o link original e sem foto,
e o HTML usa o painel com ícone da categoria no lugar.
"""

import base64
import html
import io
import json
import re
import time
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import quote, urljoin, urlparse
from urllib.request import Request, urlopen

from PIL import Image

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"
TIMEOUT = 15
WORKERS = 6
ATTEMPTS = 3
EMBED_MAX_WIDTH = 900  # px; foto reduzida para o HTML não ficar pesado (~60-120 KB cada)
MAX_IMAGE_BYTES = 20_000_000
# Imagens genéricas do site (logo, marca, imagem padrão) não servem de foto da matéria.
GENERIC_IMAGE_TERMS = ("logo", "marca", "default", "placeholder", "favicon", "share-image", "og-image-padrao")
IMAGE_PATTERNS = (
    r'<meta[^>]+property=["\']og:image(?::url)?["\'][^>]+content=["\']([^"\']+)',
    r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']og:image(?::url)?["\']',
    r'<meta[^>]+name=["\']twitter:image(?::src)?["\'][^>]+content=["\']([^"\']+)',
    r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+name=["\']twitter:image(?::src)?["\']',
)


def _fetch(url: str, data: bytes | None = None, headers: dict | None = None) -> tuple[str, str]:
    request_headers = {"User-Agent": USER_AGENT, "Accept-Language": "pt-BR,pt;q=0.9"}
    request_headers.update(headers or {})
    with urlopen(Request(url, data=data, headers=request_headers), timeout=TIMEOUT) as response:
        # A página do Google passa de 600 KB e a assinatura do link fica no fim; lê até 5 MB.
        return response.geturl(), response.read(5_000_000).decode("utf-8", "ignore")


def resolve_google_link(link: str) -> str:
    """Converte news.google.com/rss/articles/<id> no endereço real da matéria."""
    if "news.google.com" not in link:
        return link
    article_id = urlparse(link).path.rsplit("/", 1)[-1]
    _, page = _fetch(f"https://news.google.com/rss/articles/{article_id}")
    signature = re.search(r'data-n-a-sg="([^"]+)"', page).group(1)
    timestamp = re.search(r'data-n-a-ts="([^"]+)"', page).group(1)
    inner = json.dumps([
        "garturlreq",
        [["X", "X", ["X", "X"], None, None, 1, 1, "US:en", None, 1, None, None, None, None, None, 0, 1],
         "X", "X", 1, [1, 1, 1], 1, 1, None, 0, 0, None, 0],
        article_id, int(timestamp), signature,
    ])
    payload = "f.req=" + quote(json.dumps([[["Fbv4je", inner, None, "generic"]]]))
    _, response = _fetch(
        "https://news.google.com/_/DotsSplashUi/data/batchexecute",
        payload.encode(),
        {"Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"},
    )
    body = json.loads(response.split("\n\n", 1)[1])
    return json.loads(body[0][2])[1]


def page_image(url: str) -> str:
    """Foto de destaque publicada pelo site (og:image), ou "" se não houver uma boa."""
    final_url, page = _fetch(url)
    for pattern in IMAGE_PATTERNS:
        match = re.search(pattern, page, re.IGNORECASE)
        if match:
            image = urljoin(final_url, html.unescape(match.group(1).strip()))
            if image.startswith("http") and not any(term in image.lower() for term in GENERIC_IMAGE_TERMS):
                return image
    return ""


def _retry(func, arg):
    """Tenta até ATTEMPTS vezes: com várias buscas em paralelo, sites e o Google às vezes recusam uma chamada."""
    for attempt in range(ATTEMPTS):
        try:
            return func(arg)
        except Exception:
            if attempt == ATTEMPTS - 1:
                raise
            time.sleep(1.5 * (attempt + 1))


def enrich(link: str) -> tuple[str, str]:
    """(endereço real da matéria, foto). Em caso de erro devolve (link original, "")."""
    try:
        real_url = _retry(resolve_google_link, link)
    except Exception:
        return link, ""
    try:
        return real_url, _retry(page_image, real_url)
    except Exception:
        return real_url, ""


def enrich_many(links: list[str]) -> list[tuple[str, str]]:
    """enrich() em paralelo, mantendo a ordem dos links."""
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        return list(pool.map(enrich, links))


def _download_embedded(url: str) -> str:
    request = Request(url, headers={"User-Agent": USER_AGENT, "Accept": "image/avif,image/webp,image/*"})
    with urlopen(request, timeout=TIMEOUT) as response:
        raw = response.read(MAX_IMAGE_BYTES)
    image = Image.open(io.BytesIO(raw))
    image = image.convert("RGB")
    image.thumbnail((EMBED_MAX_WIDTH, EMBED_MAX_WIDTH))
    buffer = io.BytesIO()
    image.save(buffer, format="JPEG", quality=78, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(buffer.getvalue()).decode("ascii")


def embed_image(url: str) -> str:
    """Baixa a foto e devolve uma versão reduzida embutível no HTML (data URI), ou "" se falhar.

    Embutir evita que sites que bloqueiam exibição fora deles (ex.: Safra) deixem a foto
    em branco, e faz o HTML funcionar offline e anexado em e-mail.
    """
    if not url or url.startswith("data:"):
        return url or ""
    try:
        return _retry(_download_embedded, url)
    except Exception:
        return ""


def embed_many(urls: list[str]) -> list[str]:
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        return list(pool.map(embed_image, urls))

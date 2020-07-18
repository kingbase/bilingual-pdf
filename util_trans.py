import json
import requests

import docx
from tqdm import tqdm

from util_main import log, batch, para2text

TRANS_CAIYUN_URL = 'https://api.interpreter.caiyunai.com/v1/page/translator'
TRANS_CAIYUN_HEADER = {
    'Connection': 'keep-alive',
    'Origin': 'chrome-extension://jmpepeebcbihafjjadogphmbgiffiajh',
    'X-Authorization': 'token lqkr1tfixq1wa9kmj9po',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36',
    'content-type': 'application/json'
}
def translate_caiyun(texts):
    assert isinstance(texts, list)
    data = {
        'source': texts,
        'trans_type': 'en2zh',
        'request_id': 'web-translate',
        'url': 'https://www.reddit.com/forum/viewthread.php?tid=1461954',
        'page_id': 13606697,
        'replaced': True,
        'cached': True
    }
    req = requests.post(TRANS_CAIYUN_URL, headers=TRANS_CAIYUN_HEADER, data=json.dumps(data))
    targets = req.json()['target']
    assert len(texts) == len(targets)
    translated = [_['target'] for _ in targets]
    return translated

def translate_doc(doc_raw, doc_trans, trans_cache):
    doc = docx.Document(doc_raw)
    process_paragraphs = doc.paragraphs

    log.info("Translating from caiyun...")
    for batch_par in tqdm(batch(process_paragraphs, size=100)):
        source = []
        for par in batch_par:
            en_text = para2text(par)
            if len(en_text) == 0:
                continue
            if en_text in trans_cache:
                continue
            source.append(en_text)
        if len(source) == 0:
            continue
        translated = translate_caiyun(source)
        for src,target in zip(source, translated):
            trans_cache[src] = target

    log.info("Translate local...")
    for par in tqdm(process_paragraphs):
        en_text = para2text(par)
        if len(en_text) == 0:
            continue
        if en_text in trans_cache:
            par.text = trans_cache[en_text]

    log.info("Translate done and save.")
    doc.save(doc_trans)

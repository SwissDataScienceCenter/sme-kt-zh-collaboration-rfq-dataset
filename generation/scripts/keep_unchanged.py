"""Restore generated files whose content did not change since the last commit (DL-180).

Rebuilds write new timestamps, PDF IDs, mail Message-IDs and MIME boundaries even when the
content is identical. This script compares each modified tracked file with its committed
version after removing such metadata and restores the committed file if nothing else differs.
Run it after build_all.sh and build_case.py, before git add.
"""
import email, hashlib, io, subprocess, sys, zipfile
from email import policy
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SKIP_ZIP = {"docProps/core.xml", "docProps/app.xml"}

def norm(name, data):
    ext = Path(name).suffix.lower()
    if ext in (".docx", ".xlsx"):
        z = zipfile.ZipFile(io.BytesIO(data))
        return [(n, norm(n, z.read(n))) for n in sorted(z.namelist()) if n not in SKIP_ZIP]
    if ext == ".pdf":
        import pikepdf
        pdf = pikepdf.open(io.BytesIO(data)); out = []
        for p in pdf.pages:
            out.append(hashlib.sha1(b"".join(c.read_bytes() for c in (p.Contents if isinstance(p.Contents, pikepdf.Array) else [p.Contents]))).hexdigest())
            xo = p.Resources.get("/XObject", {})
            for k in sorted(xo.keys()): out.append(hashlib.sha1(xo[k].read_raw_bytes()).hexdigest())
        return out
    if ext == ".eml":
        m = email.message_from_bytes(data, policy=policy.default); out = []
        for k, v in m.items():
            if k.lower() not in ("message-id", "content-type", "in-reply-to", "references"): out.append((k, str(v)))
        for part in m.walk():
            if part.is_multipart(): continue
            fn = part.get_filename() or "body" + (part.get_content_subtype() or "")
            out.append((fn, norm(fn, part.get_payload(decode=True) or b"")))
        return out
    return hashlib.sha1(data).hexdigest()

def main():
    files = subprocess.run(["git", "diff", "--name-only", "-z"], cwd=ROOT, capture_output=True).stdout.decode().split("\0")
    restored, changed = [], []
    for f in filter(None, files):
        if not f.startswith(("company-world/", "cases/")) or f.endswith(".yaml"): continue
        p = ROOT / f
        if not p.exists(): continue
        old = subprocess.run(["git", "show", f"HEAD:{f}"], cwd=ROOT, capture_output=True).stdout
        try: same = norm(f, old) == norm(f, p.read_bytes())
        except Exception as e: same = False; print("compare failed", f, e)
        if same:
            p.write_bytes(old); restored.append(f)   # overwrite in place, no unlink needed
        else: changed.append(f)
    print(f"restored {len(restored)} files with unchanged content")
    for f in changed: print("changed:", f)

if __name__ == "__main__":
    main()

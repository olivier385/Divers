#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Archivage SPS depuis un export .pst (lecture seule du .pst, la boîte mail n'est pas touchée).
Mêmes règles que la macro ArchivageSPS.bas v1.9.
Usage :
  pst_archive.py tree  PST DEBUT FIN
  pst_archive.py run   PST DEBUT FIN [--apply]      (reprend là où il s'était arrêté, ~150 s par appel)
  pst_archive.py report PST DEBUT FIN [--apply]
Jamais d'écrasement : un fichier existant n'est jamais remplacé.
"""
import sys, os, re, json, time, pickle, datetime, unicodedata, email.utils
from email.message import EmailMessage
import pypff

HOME = os.path.expanduser("~")
MNT = os.path.join(HOME, "mnt")
JOURNAL_DIR = os.path.join(MNT, "_ARCHIVAGE MAILS")
WORK = os.path.join(HOME, "pstrun")
BOITE = "secretariat.sps@degouy.fr"
MAX_CHEMIN = 250
BUDGET = 100
COPIER_PHOTOS_MULTI_AFFAIRES = False   # False = photos d'un mail multi-affaires signalées, non copiées

CORRESPONDANCES = ("54 RUE DE ROMAINVILLE=7.2025.041;STEP CHAUMES=7.2023.463;PONT D'IVRY=7.2025.397;"
    "CHUGPN=7.2022.198;BONDOUFLE=7.2021.055;PONT AMAR=7.2021.455;hbarchitectes.fr=7.2021.455;"
    "BEAUDELAIRE=7.2026.019;BAUDELAIRE=7.2026.019;BILLETTES=7.2025.042;COLLEGE HONORE DE BALZAC=7.2023.002;"
    "FOYER PARIS DUMAS=7.2025.345;BLOMET=7.2019.347;atelierboteko=7.2019.347;GLACIERE=7.2022.500;"
    "MERCOEUR=7.2023.218;wao.paris=7.2023.218;STADE NAUTIQUE=7.2026.272;BD NEY=7.2021.105;"
    "SAINT-BERNARD-DE-LA-CHAPELLE=7.2023.282;EPHE=7.2024.324;RESTO DU COEUR=7.2025.386;RESTOS DU COEUR=7.2025.386;"
    "LEG SARTROUVILLE=7.2020.533;STATION FOCH=7.2024.384;CLICHY-SOUS-BOIS=7.2022.448;"
    "HOTEL DIEU=7.2018.194;LOURCINE=7.2023.437;RUE BERTHIER=7.2020.512")

RX_AFF = re.compile(r"(^|[^0-9])7[ ._-]{0,3}(20[0-9]{2})[ ._-]{1,3}([0-9]{3})(?![0-9])", re.I)

# ------------------------------------------------------------------ outils
def paris(dt):
    """UTC naïf -> heure de Paris (règles UE)."""
    if dt is None:
        return None
    y = dt.year
    def last_sunday(m):
        d = datetime.datetime(y, m, 31 if m in (3, 10) else 30)
        return d - datetime.timedelta(days=(d.weekday() + 1) % 7)
    start = last_sunday(3).replace(hour=1)
    end = last_sunday(10).replace(hour=1)
    off = 2 if start <= dt < end else 1
    return dt + datetime.timedelta(hours=off)

def norm(s):
    s = (s or "").upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    s = s.replace("Œ", "OE").replace("Æ", "AE")
    for c in "_.-'’":
        s = s.replace(c, " ")
    return s

def mot(u, mots):
    return re.search(r"(^|[^A-Z0-9])(" + mots + r")([^A-Z0-9]|$)", u, re.I) is not None

def affaires(texte, premiere=False):
    res = []
    for m in RX_AFF.finditer(texte or ""):
        y = int(m.group(2))
        if 2014 <= y <= datetime.date.today().year:
            n = "7.%s.%s" % (m.group(2), m.group(3))
            if n not in res:
                res.append(n)
            if premiere:
                break
    return res

def mots_seuls(s):
    return re.sub(r"[^A-Z0-9]+", " ", norm(s)).strip()

def par_correspondance(texte):
    """Recherche en mots entiers : « EPHE » ne se déclenche pas sur « STEPHEN »."""
    t = " " + mots_seuls(texte) + " "
    for paire in CORRESPONDANCES.split(";"):
        if "=" in paire:
            k, v = paire.split("=", 1)
            k = mots_seuls(k)
            if k and " " + k + " " in t:
                return [v.strip()]
    return []

def nettoyer(s):
    for c in '\\/:*?"<>|\t\r\n':
        s = s.replace(c, "_")
    s = re.sub(r" {2,}", " ", s).strip().rstrip(".")
    return s or "sans objet"

def tronquer(dossier, base, ext):
    maxi = MAX_CHEMIN - len(win(dossier)) - 1 - len(ext)
    maxi = max(maxi, 20)
    if len(base) > maxi:
        base = base[:maxi].strip()
    # Windows retire les points et espaces en fin de nom : le nom testé doit être le nom écrit
    base = base.rstrip(". ") or "sans nom"
    return base + ext

def win(p):
    """chemin local -> chemin Windows affiché (U:\\...)"""
    if p.startswith(MNT + "/"):
        r = p[len(MNT) + 1:]
        if r.startswith("_ARCHIVAGE MAILS"):
            r = "S.P.S. 26/" + r
        return "U:\\" + r.replace("/", "\\")
    return p

def cle_fichier(nom):
    base, ext = os.path.splitext(nom)
    s = base.lower()
    s = re.sub(r"^(20[0-9]{2}[ ._-]?[0-9]{2}[ ._-]?[0-9]{2})[ _-]*", "", s)
    s = re.sub(r"\([0-9]+\)\s*$", "", s)
    s = re.sub(r"[^a-z0-9]", "", s)
    return s + ext.lower()

# ------------------------------------------------------------------ serveur
class Serveur:
    def __init__(self, cache_file):
        self.cache_file = cache_file
        self.aff = {}
        self.index = {}
        self.dirty = False
        if os.path.exists(cache_file):
            with open(cache_file, "rb") as f:
                self.aff, self.index = pickle.load(f)

    def save(self):
        with open(self.cache_file, "wb") as f:
            pickle.dump((self.aff, self.index), f)

    def dossier_affaire(self, num):
        if num in self.aff:
            return self.aff[num]
        base = os.path.join(MNT, "S.P.S. " + num[4:6])
        res = ""
        try:
            for d in sorted(os.listdir(base)):
                if d.startswith(num):
                    suite = d[len(num):len(num) + 1]
                    if (suite == "" or not suite.isdigit()) and os.path.isdir(os.path.join(base, d)):
                        res = os.path.join(base, d)
                        break
        except OSError:
            pass
        self.aff[num] = res
        return res

    @staticmethod
    def sous_dossier(chemin, prefixe, defaut):
        try:
            for d in sorted(os.listdir(chemin)):
                if d.startswith(prefixe) and os.path.isdir(os.path.join(chemin, d)):
                    return os.path.join(chemin, d)
        except OSError:
            pass
        return os.path.join(chemin, defaut) if defaut else ""

    @staticmethod
    def nomme(base, nom):
        p = os.path.join(base, nom)
        return p if os.path.isdir(p) else base

    @staticmethod
    def annee(base, dt):
        p = os.path.join(base, str(dt.year))
        return p if os.path.isdir(p) else base

    def idx(self, dossier):
        if dossier in self.index:
            return self.index[dossier]
        s = {"noms": set(), "cles": {}}
        for root, dirs, files in os.walk(dossier):
            for fn in files:
                s["noms"].add(fn.lower())
                s["cles"].setdefault(cle_fichier(fn), []).append(os.path.join(root, fn))
        self.index[dossier] = s
        self.dirty = True
        return s

    def ajouter(self, dossier, nom, taille):
        s = self.idx(dossier)
        s["noms"].add(nom.lower())
        s["cles"].setdefault(cle_fichier(nom), []).append(("VIRTUEL", taille))

    def deja(self, dossier, nom, taille):
        s = self.idx(dossier)
        if nom.lower() in s["noms"]:
            return True
        for p in s["cles"].get(cle_fichier(nom), []):
            if isinstance(p, tuple):
                if p[1] == taille: return True
                continue
            try:
                if os.path.getsize(p) == taille: return True
            except OSError:
                pass
        return False

def racine_categorie(chemin, cible):
    reste = os.path.relpath(cible, chemin)
    return os.path.join(chemin, reste.split(os.sep)[0])

def contient_mot_cle(u):
    return mot(u, "PPSPS|PIC|PLAN D INSTALLATION DE CHANTIER|DIUO|CISSCT|PGC|IC|ICMOD|VIC|ICP|INSPECTION|RJ|VI|VI[0-9]+|REGISTRE|VISITE|OUVERTURE|CSPS|CR SPS|CR[0-9]*|COMPTE RENDU|COMPTERENDU|PV DE REUNION|DCE|APD|APS|AVP|LIVRET|AVIS")

def destination(S, chemin, nom, objet, externe, dt):
    u = " " + norm(nom) + " "
    if not contient_mot_cle(u):
        u = u + " " + norm(objet) + " "
    sd = S.sous_dossier
    def div():
        d = sd(chemin, "05", "")
        return d or sd(chemin, "05", "05_DIVERS")
    def rj(sub):
        return S.annee(S.nomme(sd(chemin, "01", "01_RJ et IC"), sub), dt)
    if mot(u, "PPSPS|PIC|PLAN D INSTALLATION DE CHANTIER"): return "PPSPS / PIC", sd(chemin, "02", "02_PPSPS et PIC")
    if mot(u, "AVIS") and mot(u, "PRO"): return "Avis PRO", sd(chemin, "12", "12_PRO")
    if mot(u, "AVIS") and mot(u, "DCE"): return "Avis DCE", sd(chemin, "13", "13_DCE")
    if mot(u, "AVIS") and mot(u, "APD"): return "Avis APD", sd(chemin, "11", "11_APD")
    if mot(u, "AVIS") and mot(u, "APS|AVP"): return "Avis APS / AVP", sd(chemin, "10", "10_APS et AVP")
    if mot(u, "DIUO"): return "DIUO", sd(chemin, "07", "07_DIUO")
    if mot(u, "CISSCT"): return "CISSCT", sd(chemin, "04", "04_CISSCT")
    if mot(u, "PGC"): return "PGC", sd(chemin, "09", "09_PGC")
    if mot(u, "LIVRET|LIVRET D ACCUEIL|LIVRET D ACCEUIL"): return "Autre document", div()
    if mot(u, "RJ|REGISTRE"): return "RJ", rj("RJ")
    if mot(u, "IC|ICMOD|VIC|ICP|INSPECTION COMMUNE|INSPECTION"): return "IC", rj("IC")
    if externe and mot(u, "CR[0-9]*|COMPTE RENDU|COMPTERENDU|PV DE REUNION") and not mot(u, "SPS|CSPS"):
        return "CR maîtrise d'œuvre", sd(chemin, "03", "03_CR CHANTIER")
    if mot(u, "VI|VI[0-9]+|VISITE|CR SPS|CSPS|OUVERTURE|CR[0-9]*"): return "RJ", rj("RJ")
    if mot(u, "DCE"): return "DCE", sd(chemin, "13", "13_DCE")
    if mot(u, "APD"): return "APD", sd(chemin, "11", "11_APD")
    if mot(u, "APS|AVP"): return "APS / AVP", sd(chemin, "10", "10_APS et AVP")
    return "Autre document", div()

# ------------------------------------------------------------------ pst
def props(item):
    d = {}
    try:
        for i in range(item.number_of_record_sets):
            rs = item.get_record_set(i)
            for j in range(rs.number_of_entries):
                e = rs.get_entry(j)
                t = e.entry_type
                if t in d:
                    continue
                try:
                    vt = e.value_type
                    if vt in (0x001F, 0x001E):
                        d[t] = e.get_data_as_string()
                    elif vt in (0x0003, 0x000B, 0x0014):
                        d[t] = e.get_data_as_integer()
                    elif vt == 0x0040:
                        d[t] = e.get_data_as_datetime()
                except Exception:
                    pass
    except Exception:
        pass
    return d

def ignore_mail(objet, exp):
    e, o = exp.lower(), objet.lower()
    for v in ("noreply", "no-reply", "mezzoteam", "resolving.com", "microsoftexchange", "wetransfer", "mailjet", "anthropic", "postmaster", "mailer-daemon"):
        if v in e: return True
    for v in ("réponse automatique", "reponse automatique", "automatic reply", "non remis", "undeliverable", "absente entre", "absent du"):
        if v in o: return True
    return False

def expediteur(p, headers):
    a = p.get(0x5D01) or p.get(0x5D02) or ""
    if not a and headers:
        m = re.search(r"^From:.*?<?([\w.+'-]+@[\w.-]+)>?", headers, re.M | re.I)
        if m: a = m.group(1)
    if not a:
        a = p.get(0x0C1F) or p.get(0x0065) or ""
    return a.strip()

def walk(folder, path, out):
    name = folder.name or ""
    p = path + "/" + name if path else (name or "(racine)")
    out.append((p, folder))
    for i in range(folder.number_of_sub_folders):
        walk(folder.get_sub_folder(i), p, out)

# noms exacts des dossiers système (un dossier d'assistante nommé « Notes chantier » n'est pas exclu)
EXCL = ("calendrier", "calendar", "contacts", "tâches", "tasks", "notes", "journal", "brouillons", "drafts",
        "courrier indésirable", "junk e-mail", "junk email", "historique des conversations", "conversation history",
        "flux rss", "rss feeds", "rss subscriptions", "yammer root", "suggested contacts", "contacts suggérés")
# dossiers système dont le nom a un complément (« Problèmes de synchronisation (ce poste uniquement) »...)
EXCL_DEBUT = ("problèmes de synchronisation", "sync issues", "conflits", "conflicts", "défaillances", "local failures",
              "server failures", "échecs")
ENVOYES = ("éléments envoyés", "elements envoyes", "sent items", "sent")

def dossier_mail(path, folder):
    last = path.split("/")[-1].lower()
    if last in EXCL or any(last.startswith(x) for x in EXCL_DEBUT):
        return False
    cc = props(folder).get(0x3613) or ""
    return (cc == "" or cc.startswith("IPF.Note"))

def pj_infos(att):
    p = props(att)
    nom = p.get(0x3707) or p.get(0x3704) or p.get(0x3001) or ""
    methode = p.get(0x3705, 1)
    cache = bool(p.get(0x7FFE, 0))
    ext = os.path.splitext(nom.lower())[1]
    if ext in (".png", ".jpg", ".jpeg", ".gif", ".bmp") and (p.get(0x3712) or (p.get(0x3714, 0) & 4)):
        cache = True
    return nom, methode, cache

def pj_ignoree(nom, methode, cache, taille):
    if methode in (5, 6) or cache: return True
    n = nom.lower(); ext = os.path.splitext(n)[1].lstrip(".")
    if not n: return True
    if ext in ("ics", "vcf", "p7s") or n.startswith("att0"): return True
    if ext in ("png", "jpg", "jpeg", "gif", "bmp", "emz", "wmz"):
        if n.startswith("image") or n.startswith("outlook") or n.startswith("logo") or taille < 60000: return True
    return False

def construire_eml(m, p, exp, dt_utc, pjs):
    em = EmailMessage()
    em["From"] = exp or (m.sender_name or "")
    if p.get(0x0E04): em["To"] = p.get(0x0E04).replace("\x00", "")
    if p.get(0x0E03): em["Cc"] = p.get(0x0E03).replace("\x00", "")
    em["Subject"] = (m.subject or "").replace("\r", " ").replace("\n", " ")
    if dt_utc: em["Date"] = email.utils.format_datetime(dt_utc.replace(tzinfo=datetime.timezone.utc))
    em["X-Archivage-SPS"] = "copie issue de l'export PST de secretariat.sps"
    txt = m.plain_text_body
    html = m.html_body
    txt = txt.decode("utf-8", "replace") if isinstance(txt, bytes) else (txt or "")
    em.set_content(txt or " ")
    if html:
        h = html.decode("utf-8", "replace") if isinstance(html, bytes) else html
        em.add_alternative(h, subtype="html")
    for nom, data in pjs:
        em.add_attachment(data, maintype="application", subtype="octet-stream", filename=nom)
    return em.as_bytes()

def ecrire_sans_ecraser(chemin, data):
    if os.path.exists(chemin):
        return False
    os.makedirs(os.path.dirname(chemin), exist_ok=True)
    try:
        with open(chemin, "xb") as f:
            f.write(data)
    except FileExistsError:
        return False
    return True

def est_photo(nom):
    return os.path.splitext(nom.lower())[1] in (".heic", ".heif", ".jpg", ".jpeg", ".png")

# ------------------------------------------------------------------ traitement
def traiter(pst_path, debut, fin, apply):
    os.makedirs(WORK, exist_ok=True)
    tag = "%s_%s_%s" % (debut, fin, "REEL" if apply else "SIMU")
    st_file = os.path.join(WORK, "etat_%s.json" % tag)
    rows_file = os.path.join(WORK, "lignes_%s.jsonl" % tag)
    S = Serveur(os.path.join(WORK, "cache_%s.pkl" % tag))
    etat = json.load(open(st_file)) if os.path.exists(st_file) else {"dossier": 0, "msg": 0, "fini": False, "offset": 0}
    if os.path.exists(rows_file):
        with open(rows_file, "r+b") as rf:
            rf.truncate(etat.get("offset", 0))
    if etat["fini"]:
        print("déjà terminé"); return True
    d0 = datetime.datetime.strptime(debut, "%Y-%m-%d"); d1 = datetime.datetime.strptime(fin, "%Y-%m-%d")
    t0 = time.time()
    f = pypff.file(); f.open(pst_path)
    dossiers = []; walk(f.get_root_folder(), "", dossiers)
    out = open(rows_file, "a", encoding="utf-8")
    out.seek(0, 2)
    def J(**k):
        out.write(json.dumps(k, ensure_ascii=False) + "\n")
    fi = etat["dossier"]
    while fi < len(dossiers):
        path, folder = dossiers[fi]
        if etat["msg"] == 0:
            J(type="dossier", dossier=path, n=folder.number_of_sub_messages, mail=dossier_mail(path, folder))
        if not dossier_mail(path, folder):
            fi += 1; etat.update(dossier=fi, msg=0); continue
        # dossier Éléments envoyés ou l'un de ses sous-dossiers (nom exact, pas « Absents » ni « Présentations »)
        envoye = any(c.lower() in ENVOYES for c in path.split("/"))
        n = folder.number_of_sub_messages
        mi = etat["msg"]
        while mi < n:
            if time.time() - t0 > BUDGET:
                out.flush(); etat.update(dossier=fi, msg=mi, offset=out.tell()); json.dump(etat, open(st_file, "w")); out.close(); S.save(); f.close()
                print("PAUSE dossier %d/%d msg %d/%d" % (fi, len(dossiers), mi, n)); return False
            try:
                m = folder.get_sub_message(mi)
            except Exception as e:
                J(type="erreur", dossier=path, info="lecture message %d : %s" % (mi, e)); mi += 1; continue
            mi += 1
            out.flush()
            etat.update(dossier=fi, msg=mi - 1, offset=out.tell())
            json.dump(etat, open(st_file, "w"))
            if S.dirty:
                S.save(); S.dirty = False
            try:
                dt_utc = m.delivery_time or m.client_submit_time
            except Exception:
                dt_utc = None
            dt = paris(dt_utc)
            if dt is None or not (d0 <= dt < d1):
                continue
            try:
                traiter_mail(S, m, dt, dt_utc, path, envoye, apply, J)
            except Exception as e:
                J(type="ligne", date=dt.strftime("%d/%m/%Y %H:%M"), sens="", exp="", objet=m.subject or "", affaire="", chemin="",
                  element="", fichier="", dest="", statut="ERREUR : %s" % e, dossier=path)
        fi += 1; out.flush(); etat.update(dossier=fi, msg=0, offset=out.tell()); json.dump(etat, open(st_file, "w"))
    etat["fini"] = True
    json.dump(etat, open(st_file, "w")); out.close(); S.save(); f.close()
    print("TERMINÉ"); return True

def traiter_mail(S, m, dt, dt_utc, path, envoye, apply, J):
    p = props(m)
    classe = p.get(0x001A) or "IPM.Note"
    if not classe.upper().startswith("IPM.NOTE"):
        return
    objet = m.subject or ""
    headers = m.transport_headers or ""
    exp = expediteur(p, headers)
    if ignore_mail(objet, exp):
        return
    sens = "Envoyés" if envoye else "Reçus"
    if exp.lower() == BOITE: sens = "Envoyés"
    interne = "@degouy.fr" in exp.lower() or exp.lower().startswith("/o=")
    externe = not interne
    if not externe and re.match(r"^\s*(tr|fw|fwd)\s*:", objet, re.I): externe = True
    base = dict(type="ligne", date=dt.strftime("%d/%m/%Y %H:%M"), sens=sens, exp=exp, objet=objet, dossier=path)
    affs = affaires(objet) or par_correspondance(objet + " " + exp)
    if not affs:
        body = m.plain_text_body
        body = body.decode("utf-8", "replace") if isinstance(body, bytes) else (body or "")
        affs = affaires(body[:3000], True)
    J(type="mail")
    if not affs:
        J(**base, affaire="", chemin="", element="Mail", fichier="", dest="", statut="SANS N° AFFAIRE - à classer à la main"); return
    # pièces jointes lues une seule fois
    pjs = []
    for i in range(m.number_of_attachments):
        a = m.get_attachment(i)
        nom, meth, cache = pj_infos(a)
        taille = a.size or 0
        if meth == 5 and not cache:
            J(**base, affaire=affs[0], chemin="", element="Mail joint", fichier=nom or "(mail joint)", dest="", statut="MAIL JOINT - non extrait (à voir à la main)")
            continue
        if pj_ignoree(nom, meth, cache, taille): continue
        try:
            data = a.read_buffer(taille) if taille else b""
        except Exception:
            data = b""
        pjs.append((nom, data))
    for aff in affs:
        chemin = S.dossier_affaire(aff)
        if not chemin:
            J(**base, affaire=aff, chemin="", element="Mail", fichier="", dest="", statut="DOSSIER AFFAIRE INTROUVABLE"); continue
        d06 = S.sous_dossier(chemin, "06", "06_MAILS")
        dest = os.path.join(d06, sens)
        nom_mail = tronquer(dest, dt.strftime("%Y-%m-%d %H%M") + " - " + nettoyer(objet), ".eml")
        base_mail = nom_mail[:-4]
        ix = S.idx(d06)["noms"]
        if (base_mail + ".msg").lower() in ix or nom_mail.lower() in ix:
            J(**base, affaire=aff, chemin=win(chemin), element="Mail", fichier=nom_mail, dest=win(dest), statut="DÉJÀ PRÉSENT")
        else:
            statut = "À ENREGISTRER"
            if apply:
                ok = ecrire_sans_ecraser(os.path.join(dest, nom_mail), construire_eml(m, p, exp, dt_utc, pjs))
                statut = "ENREGISTRÉ" if ok else "DÉJÀ PRÉSENT"
            S.ajouter(d06, nom_mail, 0)
            J(**base, affaire=aff, chemin=win(chemin), element="Mail", fichier=nom_mail, dest=win(dest), statut=statut)
        for nom_brut, data in pjs:
            nom = nettoyer(nom_brut)
            ch, af = chemin, aff
            n2 = affaires(nom_brut)
            if n2 and aff not in n2:
                if n2[0] in affs: continue
                c2 = S.dossier_affaire(n2[0])
                if c2: ch, af = c2, n2[0]
            if len(affs) > 1 and not n2 and est_photo(nom) and not COPIER_PHOTOS_MULTI_AFFAIRES:
                if aff == affs[0]:          # une seule ligne de journal par photo
                    J(**base, affaire=" / ".join(affs), chemin="", element="Photo", fichier=nom, dest="",
                      statut="PHOTO - mail multi-affaires, à classer à la main")
                continue
            J(type="pj")
            element, cible = destination(S, ch, nom, objet, externe, dt)
            b, e = os.path.splitext(nom)
            nom = tronquer(cible, b, e)
            rc = racine_categorie(ch, cible)
            if S.deja(rc, nom, len(data)):
                J(**base, affaire=af, chemin=win(ch), element=element, fichier=nom, dest=win(cible), statut="DÉJÀ PRÉSENT")
                continue
            statut = "À COPIER"
            if apply:
                ok = ecrire_sans_ecraser(os.path.join(cible, nom), data)
                statut = "COPIÉ" if ok else "DÉJÀ PRÉSENT (nom identique)"
            S.ajouter(rc, nom, len(data))
            J(**base, affaire=af, chemin=win(ch), element=element, fichier=nom, dest=win(cible), statut=statut)

def rapport(debut, fin, apply):
    tag = "%s_%s_%s" % (debut, fin, "REEL" if apply else "SIMU")
    rows = [json.loads(l) for l in open(os.path.join(WORK, "lignes_%s.jsonl" % tag), encoding="utf-8")]
    nom = os.path.join(JOURNAL_DIR, "Journal archivage PST %s au %s %s.csv" % (debut, fin, "" if apply else "SIMULATION"))
    nom = nom.replace(" .csv", ".csv")
    def c(s): return '"' + str(s).replace("\r", " ").replace("\n", " ").replace('"', '""') + '"'
    lignes = ["Date;Sens;Expéditeur;Objet;Affaire;Dossier affaire;Élément;Fichier;Destination;Statut;Dossier Outlook"]
    stats = {}
    for r in rows:
        if r["type"] == "dossier":
            lignes.append(";;;;;;;;;%s;%s" % (c("DOSSIER (%s éléments)%s" % (r["n"], "" if r["mail"] else " - exclu")), c(r["dossier"])))
        elif r["type"] == "ligne":
            lignes.append(";".join(c(r[k]) for k in ("date", "sens", "exp", "objet", "affaire", "chemin", "element", "fichier", "dest", "statut", "dossier")))
            k = (r["element"] if r["element"] != "Mail" else "Mail", r["statut"])
            stats[k] = stats.get(k, 0) + 1
        elif r["type"] == "erreur":
            lignes.append(";;;;;;;;;%s;%s" % (c("ERREUR " + r["info"]), c(r["dossier"])))
    base, i = nom[:-4], 2
    if os.path.exists(nom):
        nom = base + time.strftime(" %H%M") + ".csv"
    while True:                          # jamais d'écrasement : « (2) », « (3) »... si le nom est pris
        try:
            with open(nom, "x", encoding="utf-8-sig", newline="") as f:
                f.write("\r\n".join(lignes) + "\r\n")
            break
        except FileExistsError:
            nom = "%s%s (%d).csv" % (base, time.strftime(" %H%M"), i); i += 1
    print("Journal :", win(nom))
    print("mails:", sum(1 for r in rows if r["type"] == "mail"), "pj examinées:", sum(1 for r in rows if r["type"] == "pj"))
    for k in sorted(stats): print(k, stats[k])

def arbre(pst_path, debut, fin):
    f = pypff.file(); f.open(pst_path)
    d = []; walk(f.get_root_folder(), "", d)
    for p, fo in d:
        print("%6d  %s %s" % (fo.number_of_sub_messages, p, "" if dossier_mail(p, fo) else "[exclu]"))
    f.close()

if __name__ == "__main__":
    cmd, pst, debut, fin = sys.argv[1:5]
    apply = "--apply" in sys.argv
    if cmd == "tree": arbre(pst, debut, fin)
    elif cmd == "run": traiter(pst, debut, fin, apply)
    elif cmd == "report": rapport(debut, fin, apply)

# -*- coding: utf-8 -*-
"""
Table3 - ENSAM CASA (plan Table3.pdf, A2)
Equivalent Python du module VBA Excel "Table3_CATIA" : pilote CATIA V5 par COM.

Prerequis : Windows + CATIA V5 installe + Python 64 bits + pywin32
    pip install pywin32

Lancer :  python table3_catia.py

Repere : X = longueur (1350), Y = profondeur (860), Z = hauteur (2130)
Origine : coin arriere-gauche au sol. Unites : mm
"""
import math
import sys

import win32com.client

# ---------------- PARAMETRES (modifiables) -----------------------------------
LG = 1350       # longueur table et cadre haut
PR = 860        # profondeur table
H_PLAT = 880    # hauteur dessus plateau
E_PLAT = 30     # epaisseur plateau
SEC = 50        # section tubes 50x50
PAROI = 3       # epaisseur paroi tubes (0 = tubes pleins)
Z_RAIL = 90     # hauteur bas traverses basses
H_TOT = 2130    # hauteur totale montants arriere
P_SUP = 450     # profondeur cadre superieur
Y_TRAV = 405    # traverse centrale basse : Y de depart = (PR-SEC)/2

ALLEGER_PLATEAU = True   # True : plateau = peau + couronne
E_PEAU = 6               # epaisseur peau du plateau
L_COURON = 40            # largeur couronne perimetrique

# Etagere inclinee haute (vue de cote)
ANG_ETAG = 80            # angle avec le montant (=> 10 deg sous l'horizontale)
Y_ETAG_FIN = 345         # profondeur jusqu'au rebord
H_ETAG = 370             # hauteur sous rebord au-dessus du plateau
E_ETAG = 20              # epaisseur etagere
H_REBORD = 40            # hauteur totale du rebord
L_REBORD = 20            # epaisseur du rebord

# Tablette basse inclinee a 20 deg
ANG_TAB = 20
L_TAB = 330              # longueur dans le plan de coupe
E_TAB = 20
W_TAB = 450              # largeur selon X
TAB_Y = 302              # arete arriere basse (lue sur le plan)
TAB_Z = 196              # face inferieure touche l'arete de la traverse

# Panneau 1000 x 700 x 20
PAN_L = 1000
PAN_P = 700
PAN_E = 20

G_LINKERR = 0            # compteur d'echecs de liaison des lignes


# ---------------- OUTILS -----------------------------------------------------
def set_work(part, obj):
    """Definit l'objet en travail (InWorkObject)."""
    try:
        part.InWorkObject = obj
        return True
    except Exception:
        return False


def join_start(ln_b, ln_a):
    """ln_b.StartPoint = ln_a.EndPoint (chainage des lignes)."""
    global G_LINKERR
    try:
        ln_b.StartPoint = ln_a.EndPoint
    except Exception:
        G_LINKERR += 1


def poly(f, pts):
    """Polygone FERME a lignes chainees. pts = [h1, v1, h2, v2, ...]"""
    n = len(pts) // 2
    lines = []
    for i in range(n):
        j = (i + 1) % n
        lines.append(f.CreateLine(float(pts[2 * i]), float(pts[2 * i + 1]),
                                  float(pts[2 * j]), float(pts[2 * j + 1])))
    for i in range(1, n):
        join_start(lines[i], lines[i - 1])
    join_start(lines[0], lines[n - 1])


def rect(f, h1, v1, h2, v2):
    poly(f, [h1, v1, h2, v1, h2, v2, h1, v2])


def tube(f, h1, v1, h2, v2):
    """Tube carre creux : rectangle exterieur + rectangle interieur."""
    rect(f, h1, v1, h2, v2)
    if PAROI > 0:
        rect(f, h1 + PAROI, v1 + PAROI, h2 - PAROI, v2 - PAROI)


def new_plane(part, hsf, gs, base_plane, d, nom):
    """Plan decale a partir d'un plan d'origine, range dans un set geometrique."""
    ref = part.CreateReferenceFromObject(base_plane)
    pl = hsf.AddNewPlaneOffset(ref, float(d), False)
    pl.Name = nom
    gs.AppendHybridShape(pl)
    return pl


def new_sketch(body, ref, nom):
    s = body.Sketches.Add(ref)
    s.Name = nom
    return s


def make_pad(sf, part, body, sk, h, nom):
    """Cree un Pad dans le Body. 2 essais ; leve une erreur detaillee si echec."""
    pad = None
    err1 = err2 = ""
    try:
        n_elem = sk.GeometricElements.Count
    except Exception:
        n_elem = "?"

    for k in (1, 2):
        set_work(part, body)
        try:
            pad = sf.AddNewPad(sk, float(h))
            if pad is not None:
                break
        except Exception as e:
            if k == 1:
                err1 = str(e)
        try:
            part.Update()
        except Exception as e:
            err2 = str(e)

    if pad is None:
        raise RuntimeError(
            f"Echec du Pad '{nom}' (sketch '{sk.Name}').\n"
            f"Erreur AddNewPad : {err1}\n"
            f"Erreur Update    : {err2}\n"
            f"Elements du sketch : {n_elem}\n"
            f"Echecs de liaison des lignes : {G_LINKERR}"
        )
    pad.Name = nom


def masquer(doc, obj):
    """Masque un objet (sans effet bloquant en cas d'echec)."""
    try:
        sel = doc.Selection
        sel.Clear()
        sel.Add(obj)
        sel.VisProperties.SetShow(1)   # 1 = noshow
        sel.Clear()
    except Exception:
        pass


def connect_catia():
    """Se connecte a CATIA ouvert, sinon le lance."""
    try:
        return win32com.client.GetActiveObject("CATIA.Application")
    except Exception:
        pass
    try:
        catia = win32com.client.Dispatch("CATIA.Application")
        catia.Visible = True
        return catia
    except Exception as e:
        raise RuntimeError("CATIA V5 est introuvable sur ce PC "
                           "(CATIA.Application).") from e


# ---------------- PROGRAMME PRINCIPAL ----------------------------------------
def table3_generer():
    catia = connect_catia()

    # --- Nouveau document Part
    doc = catia.Documents.Add("Part")
    part = doc.Part
    sf = part.ShapeFactory
    hsf = part.HybridShapeFactory
    org = part.OriginElements

    body = part.MainBody
    body.Name = "Table3"
    if not set_work(part, body):
        raise RuntimeError("Impossible de definir l'objet en travail (InWorkObject).")

    zh = H_TOT - SEC          # bas du cadre superieur (2080)

    # === Plans de reference ==================================================
    gs = part.HybridBodies.Add()
    gs.Name = "Plans_Reference"
    pl_x50 = new_plane(part, hsf, gs, org.PlaneYZ, SEC, "P_X50")
    pl_y50 = new_plane(part, hsf, gs, org.PlaneZX, SEC, "P_Y50")
    pl_x450 = new_plane(part, hsf, gs, org.PlaneYZ, (LG - W_TAB) / 2, "P_X450")
    pl_z850 = new_plane(part, hsf, gs, org.PlaneXY, H_PLAT - E_PLAT, "P_Z850")
    pl_z874 = None
    if ALLEGER_PLATEAU:
        pl_z874 = new_plane(part, hsf, gs, org.PlaneXY, H_PLAT - E_PEAU, "P_Z874")
    part.Update()
    set_work(part, body)

    r_xy = part.CreateReferenceFromObject(org.PlaneXY)
    r_yz = part.CreateReferenceFromObject(org.PlaneYZ)
    r_x50 = part.CreateReferenceFromObject(pl_x50)
    r_y50 = part.CreateReferenceFromObject(pl_y50)
    r_x450 = part.CreateReferenceFromObject(pl_x450)
    r_z850 = part.CreateReferenceFromObject(pl_z850)
    r_z874 = part.CreateReferenceFromObject(pl_z874) if ALLEGER_PLATEAU else None

    # #########################################################################
    # BODY 1 : OSSATURE TUBULAIRE
    # #########################################################################

    # 1. Montants arriere pleine hauteur
    sk = new_sketch(body, r_xy, "Sk_MontantsArriere")
    f = sk.OpenEdition()
    tube(f, 0, 0, SEC, SEC)
    tube(f, LG - SEC, 0, LG, SEC)
    sk.CloseEdition()
    make_pad(sf, part, body, sk, H_TOT, "Pad_MontantsArriere")

    # 2. Traverses basses selon Y (gauche/droite)
    sk = new_sketch(body, r_y50, "Sk_TraversesY")
    f = sk.OpenEdition()
    tube(f, Z_RAIL, 0, Z_RAIL + SEC, SEC)
    tube(f, Z_RAIL, LG - SEC, Z_RAIL + SEC, LG)
    sk.CloseEdition()
    make_pad(sf, part, body, sk, PR - 2 * SEC, "Pad_TraversesY")

    # 3. Pieds avant
    sk = new_sketch(body, r_xy, "Sk_PiedsAvant")
    f = sk.OpenEdition()
    tube(f, 0, PR - SEC, SEC, PR)
    tube(f, LG - SEC, PR - SEC, LG, PR)
    sk.CloseEdition()
    make_pad(sf, part, body, sk, H_PLAT - E_PLAT, "Pad_PiedsAvant")

    # 4. Traverse centrale basse selon X
    sk = new_sketch(body, r_x50, "Sk_TraverseCentrale")
    f = sk.OpenEdition()
    tube(f, Y_TRAV, Z_RAIL, Y_TRAV + SEC, Z_RAIL + SEC)
    sk.CloseEdition()
    make_pad(sf, part, body, sk, LG - 2 * SEC, "Pad_TraverseCentrale")

    # 5. Cadre superieur : cotes gauche/droite
    sk = new_sketch(body, r_y50, "Sk_CadreHautY")
    f = sk.OpenEdition()
    tube(f, zh, 0, H_TOT, SEC)
    tube(f, zh, LG - SEC, H_TOT, LG)
    sk.CloseEdition()
    make_pad(sf, part, body, sk, P_SUP - 2 * SEC, "Pad_CadreHautY")

    # 6. Cadre superieur : traverses arriere + avant, pleine longueur
    sk = new_sketch(body, r_yz, "Sk_CadreHautX")
    f = sk.OpenEdition()
    tube(f, 0, zh, SEC, H_TOT)
    tube(f, P_SUP - SEC, zh, P_SUP, H_TOT)
    sk.CloseEdition()
    make_pad(sf, part, body, sk, LG, "Pad_CadreHautX")

    # #########################################################################
    # BODY 2 : PLATEAU 1350 x 860 x 30 (allege : couronne + peau)
    # #########################################################################
    bd2 = part.Bodies.Add()
    bd2.Name = "Plateau"
    if not set_work(part, bd2):
        raise RuntimeError("InWorkObject impossible (Plateau).")

    if ALLEGER_PLATEAU:
        sk = new_sketch(bd2, r_z850, "Sk_Couronne")
        f = sk.OpenEdition()
        rect(f, 0, 0, LG, PR)
        rect(f, L_COURON, L_COURON, LG - L_COURON, PR - L_COURON)
        sk.CloseEdition()
        make_pad(sf, part, bd2, sk, E_PLAT - E_PEAU, "Pad_Couronne")

        sk = new_sketch(bd2, r_z874, "Sk_Peau")
        f = sk.OpenEdition()
        rect(f, 0, 0, LG, PR)
        sk.CloseEdition()
        make_pad(sf, part, bd2, sk, E_PEAU, "Pad_Peau")
    else:
        sk = new_sketch(bd2, r_z850, "Sk_Plateau")
        f = sk.OpenEdition()
        rect(f, 0, 0, LG, PR)
        sk.CloseEdition()
        make_pad(sf, part, bd2, sk, E_PLAT, "Pad_Plateau")

    # #########################################################################
    # BODY 3 : ETAGERE INCLINEE (80 deg / montant) AVEC REBORD
    # #########################################################################
    bd2 = part.Bodies.Add()
    bd2.Name = "Etagere"
    if not set_work(part, bd2):
        raise RuntimeError("InWorkObject impossible (Etagere).")

    a = math.radians(90 - ANG_ETAG)          # pente sous l'horizontale = 10 deg
    ca, sa = math.cos(a), math.sin(a)
    dY, dZ = ca, -sa                         # direction vers l'avant (descend)
    nY, nZ = sa, ca                          # normale vers le haut
    b2y, b2z = Y_ETAG_FIN, H_PLAT + H_ETAG   # bas du rebord
    b1y = SEC                                # au contact du montant
    b1z = b2z + (b2y - b1y) * math.tan(a)

    sk = new_sketch(bd2, r_x50, "Sk_Etagere")
    f = sk.OpenEdition()
    poly(f, [
        b1y, b1z,
        b2y, b2z,
        b2y + H_REBORD * nY, b2z + H_REBORD * nZ,
        b2y + H_REBORD * nY - L_REBORD * dY, b2z + H_REBORD * nZ - L_REBORD * dZ,
        b2y + E_ETAG * nY - L_REBORD * dY, b2z + E_ETAG * nZ - L_REBORD * dZ,
        b1y + E_ETAG * nY, b1z + E_ETAG * nZ,
    ])
    sk.CloseEdition()
    make_pad(sf, part, bd2, sk, LG - 2 * SEC, "Pad_Etagere")

    # #########################################################################
    # BODY 4 : TABLETTE BASSE INCLINEE A 20 DEG (330 x 450 x 20)
    # #########################################################################
    bd2 = part.Bodies.Add()
    bd2.Name = "Tablette_20deg"
    if not set_work(part, bd2):
        raise RuntimeError("InWorkObject impossible (Tablette).")

    c20, s20 = math.cos(math.radians(ANG_TAB)), math.sin(math.radians(ANG_TAB))
    sk = new_sketch(bd2, r_x450, "Sk_TabletteBasse")
    f = sk.OpenEdition()
    poly(f, [
        TAB_Y, TAB_Z,
        TAB_Y + L_TAB * c20, TAB_Z - L_TAB * s20,
        TAB_Y + L_TAB * c20 + E_TAB * s20, TAB_Z - L_TAB * s20 + E_TAB * c20,
        TAB_Y + E_TAB * s20, TAB_Z + E_TAB * c20,
    ])
    sk.CloseEdition()
    make_pad(sf, part, bd2, sk, W_TAB, "Pad_TabletteBasse")

    # #########################################################################
    # BODY 5 : PANNEAU 1000 x 700 x 20 (piece representee a part)
    # #########################################################################
    bd2 = part.Bodies.Add()
    bd2.Name = "Panneau_1000x700"
    if not set_work(part, bd2):
        raise RuntimeError("InWorkObject impossible (Panneau).")

    sk = new_sketch(bd2, r_xy, "Sk_Panneau")
    f = sk.OpenEdition()
    rect(f, (LG - PAN_L) / 2, PR + 100, (LG - PAN_L) / 2 + PAN_L, PR + 100 + PAN_P)
    sk.CloseEdition()
    make_pad(sf, part, bd2, sk, PAN_E, "Pad_Panneau")

    # === Mise a jour unique + masquage des plans + cadrage ===================
    set_work(part, body)
    part.Update()
    masquer(doc, gs)
    try:
        catia.ActiveWindow.ActiveViewer.Reframe()
    except Exception:
        pass

    print("Table3 generee dans Part Design.")
    print(f"Body 'Table3' = ossature tubulaire creuse (paroi {PAROI} mm).")
    print("Verifier la masse : Analyze > Measure Inertia (apres application du materiau).")


if __name__ == "__main__":
    try:
        table3_generer()
    except Exception as exc:
        print("\nERREUR :", exc, file=sys.stderr)
        sys.exit(1)
# -*- coding: utf-8 -*-
"""
Skarbnik – program dla skarbnika klasowego (Windows, bez MS Office).

Zakładki: Dashboard | Uczniowie i Wpłaty | Rejestr Wydatków | Słowniki i Parametry.
Każda zmiana jest od razu zapisywana na dysk (autozapis) i przelicza wszystko na nowo.

Dlaczego tkinter: jest wbudowany w Pythona, więc .exe jest jednym plikiem bez zależności
graficznych. Logika (sumy, salda, statusy) siedzi w model.py.
"""
import datetime
import os
import shutil
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import model as M

NAZWA_PLIKU = "skarbnik_dane.json"
KOL_STATUS = {"Zaległość": ("#FFC7CE", "#9C0006"), "Zapłacone": ("#E2F0D9", "#006100"),
              "Nadpłata": ("#DDEBF7", "#1F4E78")}


def domyslna_sciezka():
    """Dane leżą obok programu (łatwo skopiować jako backup). Jeśli folder jest
    tylko-do-odczytu (np. Program Files), używamy folderu użytkownika."""
    folder = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
    try:
        test = os.path.join(folder, ".test_zapisu")
        with open(test, "w") as f:
            f.write("x")
        os.remove(test)
    except OSError:
        folder = os.path.join(os.path.expanduser("~"), "Skarbnik")
        os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, NAZWA_PLIKU)


class Formularz(tk.Toplevel):
    """Uniwersalne okno z polami (tekst / lista rozwijana) i walidacją."""

    def __init__(self, master, tytul, pola, wartosci=None, waliduj=None):
        super().__init__(master)
        self.title(tytul)
        self.transient(master)
        self.resizable(False, False)
        self.wynik, self._waliduj, self._w = None, waliduj, {}
        wartosci = wartosci or {}
        for i, (klucz, etykieta, typ, opcje) in enumerate(pola):
            ttk.Label(self, text=etykieta).grid(row=i, column=0, sticky="e", padx=8, pady=4)
            if typ == "combo":
                w = ttk.Combobox(self, values=opcje, state="readonly", width=38)
                w.set(wartosci.get(klucz, opcje[0] if opcje else ""))
            else:
                w = ttk.Entry(self, width=41)
                w.insert(0, wartosci.get(klucz, ""))
            w.grid(row=i, column=1, padx=8, pady=4)
            self._w[klucz] = w
        przyciski = ttk.Frame(self)
        przyciski.grid(row=len(pola), column=0, columnspan=2, pady=8)
        ttk.Button(przyciski, text="OK", command=self._ok).pack(side="left", padx=6)
        ttk.Button(przyciski, text="Anuluj", command=self.destroy).pack(side="left", padx=6)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        next(iter(self._w.values())).focus_set()
        self.grab_set()

    def _ok(self):
        wart = {k: w.get().strip() for k, w in self._w.items()}
        if self._waliduj:
            blad = self._waliduj(wart)
            if blad:
                messagebox.showwarning("Popraw dane", blad, parent=self)
                return
        self.wynik = wart
        self.destroy()


def pytaj(master, tytul, pola, wartosci=None, waliduj=None):
    f = Formularz(master, tytul, pola, wartosci, waliduj)
    master.wait_window(f)
    return f.wynik


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Skarbnik klasowy")
        self.geometry("1280x760")
        self.minsize(1000, 600)
        st = ttk.Style(self)
        if "vista" in st.theme_names():
            st.theme_use("vista")
        st.configure("Treeview", rowheight=24)
        st.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        self.sciezka = domyslna_sciezka()
        try:
            self.dane = M.wczytaj(self.sciezka) if os.path.exists(self.sciezka) else M.nowe()
        except Exception as e:
            messagebox.showerror("Błąd odczytu", f"Nie można wczytać {self.sciezka}\n{e}\n\nStart z pustymi danymi "
                                 "(plik zostaje nietknięty do pierwszej zmiany).")
            self.dane = M.nowe()

        self._menu()
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)
        self.tab_d, self.tab_u = ttk.Frame(self.nb), ttk.Frame(self.nb)
        self.tab_r, self.tab_s = ttk.Frame(self.nb), ttk.Frame(self.nb)
        for t, n in ((self.tab_d, "Dashboard"), (self.tab_u, "Uczniowie i Wpłaty"),
                     (self.tab_r, "Rejestr Wydatków"), (self.tab_s, "Słowniki i Parametry")):
            self.nb.add(t, text=n)
        self._zbuduj_dashboard()
        self._zbuduj_uczniow()
        self._zbuduj_rejestr()
        self._zbuduj_slowniki()
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self.odswiez_wszystko())
        self.bind("<Control-s>", lambda e: self.zapisz_jako())
        self.odswiez_wszystko()

    # ------------------------------------------------------------------ menu i zapis
    def _menu(self):
        m = tk.Menu(self)
        plik = tk.Menu(m, tearoff=0)
        plik.add_command(label="Nowy (czyści dane – robi kopię)", command=self.nowy)
        plik.add_command(label="Wczytaj dane przykładowe", command=self.wczytaj_przyklad)
        plik.add_separator()
        plik.add_command(label="Otwórz plik danych (.json)…", command=self.otworz_json)
        plik.add_command(label="Zapisz jako… (Ctrl+S)", command=self.zapisz_jako)
        plik.add_separator()
        plik.add_command(label="Importuj z Excela (Skarbnik_Klasowy.xlsx)…", command=self.import_excel)
        plik.add_command(label="Eksportuj raport do Excela…", command=self.eksport_excel)
        plik.add_separator()
        plik.add_command(label="Zakończ", command=self.destroy)
        m.add_cascade(label="Plik", menu=plik)
        rap = tk.Menu(m, tearoff=0)
        rap.add_command(label="Kopiuj listę zaległości do schowka", command=self.kopiuj_zaleglosci)
        m.add_cascade(label="Raporty", menu=rap)
        pom = tk.Menu(m, tearoff=0)
        pom.add_command(label="O programie", command=lambda: messagebox.showinfo(
            "O programie", f"Skarbnik klasowy\nDane (autozapis):\n{self.sciezka}\n\n"
            "Kopia poprzedniej wersji: plik .bak obok danych."))
        m.add_cascade(label="Pomoc", menu=pom)
        self.config(menu=m)

    def zmieniono(self):
        """Wołane po KAŻDEJ zmianie: autozapis + przeliczenie widoków."""
        try:
            M.zapisz(self.dane, self.sciezka)
        except OSError as e:
            messagebox.showerror("Nie udało się zapisać", f"{e}\n\nSpróbuj 'Zapisz jako…' w innym folderze.")
        self.odswiez_wszystko()

    def kopia_zapasowa(self):
        # Przed operacjami destrukcyjnymi zostawiamy kopię z datą – cofnięcie błędu zawsze możliwe
        if os.path.exists(self.sciezka):
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(self.sciezka, f"{self.sciezka}.{ts}.bak")

    def nowy(self):
        if messagebox.askyesno("Nowy", "Wyczyścić wszystkie dane? (kopia zapasowa zostanie zapisana obok pliku)"):
            self.kopia_zapasowa()
            self.dane = M.nowe()
            self._przebuduj_kolumny()
            self.zmieniono()

    def wczytaj_przyklad(self):
        if messagebox.askyesno("Dane przykładowe", "Zastąpić obecne dane przykładowymi? (kopia zostanie zapisana)"):
            self.kopia_zapasowa()
            self.dane = M.przyklad()
            self._przebuduj_kolumny()
            self.zmieniono()

    def otworz_json(self):
        p = filedialog.askopenfilename(filetypes=[("Dane Skarbnika", "*.json")])
        if p:
            try:
                self.dane = M.wczytaj(p)
            except Exception as e:
                messagebox.showerror("Błąd", str(e)); return
            self.sciezka = p
            self._przebuduj_kolumny()
            self.odswiez_wszystko()

    def zapisz_jako(self):
        p = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Dane Skarbnika", "*.json")],
                                         initialfile=NAZWA_PLIKU)
        if p:
            self.sciezka = p
            self.zmieniono()

    def import_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not p:
            return
        try:
            nowe = M.import_xlsx(p)
        except Exception as e:
            messagebox.showerror("Import nieudany", f"To nie wygląda na plik z tego programu.\n{e}"); return
        self.kopia_zapasowa()
        self.dane = nowe
        self._przebuduj_kolumny()
        self.zmieniono()
        messagebox.showinfo("Import", f"Wczytano: uczniów {len(nowe['uczniowie'])}, wydatków {len(nowe['wydatki'])}.")

    def eksport_excel(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                         initialfile="Raport_skarbnika.xlsx")
        if p:
            try:
                M.eksport_xlsx(self.dane, p)
                messagebox.showinfo("Eksport", "Zapisano raport.")
            except Exception as e:
                messagebox.showerror("Eksport nieudany", str(e))

    def kopiuj_zaleglosci(self):
        tekst = M.lista_zaleglosci(self.dane)
        if not tekst:
            messagebox.showinfo("Zaległości", "Brak zaległości 🎉"); return
        self.clipboard_clear(); self.clipboard_append(tekst)
        messagebox.showinfo("Skopiowano", "Lista zaległości jest w schowku – wklej ją do wiadomości.\n\n" + tekst)

    def odswiez_wszystko(self):
        self.odswiez_dashboard()
        self.odswiez_uczniow()
        self.odswiez_rejestr()

    def _przebuduj_kolumny(self):
        self._kolumny_uczniow()
        self._wczytaj_slowniki_do_pol()

    # ------------------------------------------------------------------ DASHBOARD
    def _zbuduj_dashboard(self):
        f = self.tab_d
        self.lbl_tytul = ttk.Label(f, font=("Segoe UI", 16, "bold"), foreground="#1F3864")
        self.lbl_tytul.pack(anchor="w", padx=12, pady=(10, 4))
        karty = ttk.Frame(f)
        karty.pack(fill="x", padx=8)
        self.karty = {}
        for i, (k, n) in enumerate([("budzet", "Całkowity budżet"), ("wplaty", "Łączne wpłaty"),
                                    ("wydatki", "Łączne wydatki"), ("saldo", "Aktualne saldo"),
                                    ("sciagalnosc", "Ściągalność składek"), ("zaleglosci", "Zaległości")]):
            fr = tk.Frame(karty, bg="#EAF1FB", bd=1, relief="solid")
            fr.grid(row=0, column=i, padx=4, sticky="nsew")
            karty.columnconfigure(i, weight=1)
            tk.Label(fr, text=n, bg="#1F3864", fg="white", font=("Segoe UI", 9, "bold")).pack(fill="x")
            v = tk.Label(fr, text="-", bg="#EAF1FB", font=("Segoe UI", 15, "bold"))
            v.pack(pady=10)
            self.karty[k] = v

        srodek = ttk.Frame(f)
        srodek.pack(fill="x", padx=8, pady=8)
        self.t_fin = ttk.Treeview(srodek, columns=("poz", "kw"), show="headings", height=9)
        self.t_fin.heading("poz", text="Podsumowanie finansów"); self.t_fin.heading("kw", text="Kwota")
        self.t_fin.column("poz", width=290); self.t_fin.column("kw", width=140, anchor="e")
        self.t_fin.pack(side="left", padx=4)
        self.t_stat = ttk.Treeview(srodek, columns=("poz", "n"), show="headings", height=4)
        self.t_stat.heading("poz", text="Status uczniów"); self.t_stat.heading("n", text="Liczba")
        self.t_stat.column("poz", width=180); self.t_stat.column("n", width=80, anchor="e")
        self.t_stat.pack(side="left", padx=20, anchor="n")

        ttk.Label(f, text="Wydatki wg kategorii", font=("Segoe UI", 11, "bold"),
                  foreground="#1F3864").pack(anchor="w", padx=12)
        dol = ttk.Frame(f)
        dol.pack(fill="both", expand=True, padx=8, pady=4)
        self.t_kat = ttk.Treeview(dol, show="headings", height=8)
        self.t_kat.pack(side="left", anchor="n", padx=4)
        self.wykres = tk.Canvas(dol, bg="white", highlightthickness=1, highlightbackground="#BFBFBF")
        self.wykres.pack(side="left", fill="both", expand=True, padx=12)
        self.wykres.bind("<Configure>", lambda e: self._rysuj_wykres())
        self._dane_wykresu = []

    def odswiez_dashboard(self):
        p, d = M.podsumowanie(self.dane), self.dane
        par = d["parametry"]
        self.lbl_tytul.config(text=f'Skarbnik klasowy – {par["klasa"]}  {("| rok szkolny " + par["rok"]) if par["rok"] else ""}')
        for k in ("budzet", "wplaty", "wydatki", "saldo", "zaleglosci"):
            self.karty[k].config(text=M.fmt_zl(p[k]), fg="#C00000" if (k == "saldo" and p[k] < 0) or
                                 (k == "zaleglosci" and p[k] > 0) else "#000000")
        sc = p["sciagalnosc"]
        self.karty["sciagalnosc"].config(text=f"{sc*100:.1f}".replace(".", ",") + " %",
                                         fg="#548235" if sc >= 0.9 else ("#BF8F00" if sc >= 0.7 else "#C00000"))
        self.t_fin.delete(*self.t_fin.get_children())
        for n, v in [("Saldo początkowe", p["saldo_poczatkowe"]), ("Wpłaty składek od rodziców", p["wplaty"]),
                     ("Inne przychody (darowizny, odsetki)", p["inne"]),
                     ("Razem środki (z saldem początkowym)", p["razem_srodki"]),
                     ("Łączne wydatki", p["wydatki"]), ("Aktualne saldo", p["saldo"]),
                     ("Planowany budżet (składki wymagane)", p["budzet"]),
                     ("Zaległości do ściągnięcia", p["zaleglosci"])]:
            self.t_fin.insert("", "end", values=(n, M.fmt_zl(v)))
        self.t_stat.delete(*self.t_stat.get_children())
        for n, v in [("Liczba uczniów", p["liczba"]), ("Zapłacone", p["zaplacone"]),
                     ("Zaległość", p["zaleglosc"]), ("Nadpłata", p["nadplata"])]:
            self.t_stat.insert("", "end", values=(n, v), tags=(n,) if n in KOL_STATUS else ())
        for n, (bg, fg) in KOL_STATUS.items():
            self.t_stat.tag_configure(n, background=bg, foreground=fg)

        formy = d["formy_platnosci"]
        kol = ["kat"] + [f"f{i}" for i in range(len(formy))] + ["razem", "proc"]
        self.t_kat["columns"] = kol
        self.t_kat.heading("kat", text="Kategoria"); self.t_kat.column("kat", width=190)
        for i, fn in enumerate(formy):
            self.t_kat.heading(f"f{i}", text=fn); self.t_kat.column(f"f{i}", width=100, anchor="e")
        self.t_kat.heading("razem", text="Razem"); self.t_kat.column("razem", width=110, anchor="e")
        self.t_kat.heading("proc", text="% udziału"); self.t_kat.column("proc", width=80, anchor="e")
        self.t_kat.delete(*self.t_kat.get_children())
        for k in p["kategorie"]:
            self.t_kat.insert("", "end", values=[k["kategoria"]] + [M.fmt_kw(k["formy"][f]) for f in formy] +
                              [M.fmt_kw(k["razem"]), f'{k["proc"]*100:.1f}'.replace(".", ",") + " %"])
        wiersz_razem = ["RAZEM"] + [M.fmt_kw(sum(k["formy"][f] for k in p["kategorie"])) for f in formy] + \
                       [M.fmt_kw(p["wydatki"] - p["nieprzypisane"]), ""]
        self.t_kat.insert("", "end", values=wiersz_razem, tags=("razem",))
        self.t_kat.tag_configure("razem", background="#D9E1F2", font=("Segoe UI", 9, "bold"))
        if p["nieprzypisane"]:
            self.t_kat.insert("", "end", values=["⚠ Wydatki bez kategorii"] + [""] * len(formy) +
                              [M.fmt_kw(p["nieprzypisane"]), ""], tags=("blad",))
            self.t_kat.tag_configure("blad", background="#FFC7CE", foreground="#9C0006")
        self._dane_wykresu = [(k["kategoria"], k["razem"]) for k in p["kategorie"]]
        self._rysuj_wykres()

    def _rysuj_wykres(self):
        c = self.wykres
        c.delete("all")
        lista = self._dane_wykresu
        w = max(c.winfo_width(), 400)
        maks = max([k for _, k in lista] or [0]) or 1
        y = 12
        c.create_text(10, y, text="Wydatki wg kategorii (zł)", anchor="w", font=("Segoe UI", 10, "bold"))
        y += 24
        for nazwa, kw in lista:
            c.create_text(10, y + 10, text=nazwa, anchor="w", font=("Segoe UI", 9))
            dl = int(max(w - 400, 50) * kw / maks)
            c.create_rectangle(190, y, 190 + dl, y + 20, fill="#2F5597", outline="")
            c.create_text(190 + dl + 6, y + 10, text=M.fmt_zl(kw), anchor="w", font=("Segoe UI", 9))
            y += 28

    # ------------------------------------------------------------------ UCZNIOWIE
    def _zbuduj_uczniow(self):
        f = self.tab_u
        pasek = ttk.Frame(f, padding=4)
        pasek.pack(fill="x")
        ttk.Button(pasek, text="➕ Dodaj ucznia", command=self.dodaj_ucznia).pack(side="left")
        ttk.Button(pasek, text="💰 Dodaj wpłatę", command=self.dodaj_wplate).pack(side="left", padx=6)
        ttk.Button(pasek, text="🗑 Usuń zaznaczonego", command=self.usun_ucznia).pack(side="left")
        ttk.Label(pasek, foreground="#595959",
                  text="  Kliknij dwukrotnie komórkę (imię, nazwisko, wpłata, zniżka), aby ją edytować. "
                       "Enter = zapisz i przejdź niżej.").pack(side="left")
        ramka = ttk.Frame(f)
        ramka.pack(fill="both", expand=True)
        self.tree_u = ttk.Treeview(ramka, show="headings", selectmode="browse")
        v = ttk.Scrollbar(ramka, orient="vertical", command=self.tree_u.yview)
        h = ttk.Scrollbar(ramka, orient="horizontal", command=self.tree_u.xview)
        self.tree_u.configure(yscrollcommand=v.set, xscrollcommand=h.set)
        v.pack(side="right", fill="y"); h.pack(side="bottom", fill="x")
        self.tree_u.pack(fill="both", expand=True)
        self.tree_u.bind("<Double-1>", self._dbl_uczen)
        for n, (bg, fg) in KOL_STATUS.items():
            self.tree_u.tag_configure(n, background=bg, foreground=fg)
        self.tree_u.tag_configure("suma", background="#D9E1F2", font=("Segoe UI", 9, "bold"))
        self.tree_u.tag_configure("podsum", background="#F2F2F2")
        self._kolumny_uczniow()

    def _kolumny_uczniow(self):
        sk = self.dane["skladki"]
        self.kol_u = ["nr", "imie", "nazwisko"] + [f"s{i}" for i in range(len(sk))] + \
                     ["znizka", "wym", "wpl", "saldo", "status"]
        t = self.tree_u
        t["columns"] = self.kol_u
        nag = {"nr": "Nr", "imie": "Imię", "nazwisko": "Nazwisko", "znizka": "Zniżka (zł)", "wym": "Suma wymagana",
               "wpl": "Suma wpłacona", "saldo": "Saldo / Zaległość", "status": "Status"}
        for i, s in enumerate(sk):
            nag[f"s{i}"] = f'{s["nazwa"]} ({s["stawka"]:g} zł)'
        for k in self.kol_u:
            t.heading(k, text=nag[k])
            szer = {"nr": 50, "imie": 110, "nazwisko": 140, "status": 100}.get(k, 150 if k.startswith("s") else 115)
            t.column(k, width=szer, anchor="center" if k in ("nr", "status") else ("w" if k in ("imie", "nazwisko") else "e"),
                     stretch=False)

    def odswiez_uczniow(self):
        t, d = self.tree_u, self.dane
        y = t.yview()[0]
        wybrany = t.selection()
        t.delete(*t.get_children())
        for i, u in enumerate(d["uczniowie"]):
            st = M.status(u, d)
            vals = [u["nr"], u["imie"], u["nazwisko"]] + [M.fmt_wplata(x) for x in u["wplaty"]] + \
                   [M.fmt_wplata(u.get("znizka")), M.fmt_kw(M.wymagana(u, d)), M.fmt_kw(M.wplacona(u)),
                    M.fmt_kw(M.saldo(u, d)), st]
            t.insert("", "end", iid=f"u{i}", values=vals, tags=(st,))
        s, n = M.sumy_uczniow(d), len(d["skladki"])
        etyk = lambda txt: ["", txt, ""]
        t.insert("", "end", iid="t0", tags=("suma",), values=etyk("Suma wpłat") + [M.fmt_kw(x) for x in s["suma"]] +
                 [M.fmt_kw(s["znizki_suma"]), M.fmt_kw(s["wym_suma"]), M.fmt_kw(s["wpl_suma"]), M.fmt_kw(s["saldo_suma"]), ""])
        t.insert("", "end", iid="t1", tags=("podsum",), values=etyk("Liczba wpłacających") + [str(x) for x in s["liczba"]] +
                 ["", "", str(s["wpl_liczba"]), "", ""])
        t.insert("", "end", iid="t2", tags=("podsum",), values=etyk("Średnia wpłata") + [M.fmt_kw(x) for x in s["srednia"]] +
                 ["", "", M.fmt_kw(s["wpl_srednia"]), "", ""])
        for iid in wybrany:
            if t.exists(iid):
                t.selection_set(iid)
        t.yview_moveto(y)

    def _dbl_uczen(self, event):
        t = self.tree_u
        if t.identify_region(event.x, event.y) != "cell":
            return
        iid, kol = t.identify_row(event.y), int(t.identify_column(event.x)[1:]) - 1
        self._edytor_uczen(iid, kol)

    def _edytor_uczen(self, iid, kol):
        t = self.tree_u
        if not iid or not iid.startswith("u"):
            return
        klucz = self.kol_u[kol]
        if klucz not in ("nr", "imie", "nazwisko", "znizka") and not klucz.startswith("s"):
            return                                   # kolumny wyliczane są tylko do odczytu
        idx = int(iid[1:])
        u = self.dane["uczniowie"][idx]
        t.see(iid)
        t.update_idletasks()
        x, y, w, h = t.bbox(iid, f"#{kol+1}")
        if klucz.startswith("s"):
            obecna = u["wplaty"][int(klucz[1:])]
        elif klucz == "znizka":
            obecna = u.get("znizka") or 0
        else:
            obecna = u[klucz]
        tekst = "" if obecna in (0, 0.0, "") else (M.fmt_kw(obecna).replace(" ", "") if isinstance(obecna, float) else str(obecna))
        e = ttk.Entry(t)
        e.place(x=x, y=y, width=w, height=h)
        e.insert(0, tekst); e.focus_set(); e.select_range(0, "end")
        stan = {"koniec": False}

        def zatwierdz(dalej=False):
            if stan["koniec"]:
                return
            stan["koniec"] = True
            wartosc = e.get()
            e.destroy()
            try:
                if klucz == "nr":
                    u["nr"] = int(wartosc)
                elif klucz in ("imie", "nazwisko"):
                    u[klucz] = wartosc.strip()
                else:
                    kw = M.parse_kwota(wartosc)
                    if kw < 0:
                        raise ValueError("Kwota nie może być ujemna")
                    if klucz == "znizka":
                        u["znizka"] = kw
                    else:
                        u["wplaty"][int(klucz[1:])] = kw
            except ValueError as err:
                messagebox.showwarning("Błędna wartość", f"Nie rozumiem: '{wartosc}'. Wpisz liczbę (np. 150 lub 150,50).\n{err}")
                return
            self.zmieniono()
            nast = f"u{idx+1}"
            if dalej and t.exists(nast):
                self.after(30, lambda: self._edytor_uczen(nast, kol))

        e.bind("<Return>", lambda ev: zatwierdz(True))
        e.bind("<Escape>", lambda ev: (stan.update(koniec=True), e.destroy()))
        e.bind("<FocusOut>", lambda ev: zatwierdz(False))

    def dodaj_ucznia(self):
        def waliduj(w):
            return "Wpisz nazwisko ucznia." if not w["nazwisko"] else None
        nr = max([u["nr"] for u in self.dane["uczniowie"] if isinstance(u["nr"], int)] or [0]) + 1
        w = pytaj(self, "Nowy uczeń", [("nr", "Nr w dzienniku", "entry", None), ("imie", "Imię", "entry", None),
                                       ("nazwisko", "Nazwisko", "entry", None)], {"nr": str(nr)}, waliduj)
        if w:
            try:
                nr = int(w["nr"])
            except ValueError:
                pass
            self.dane["uczniowie"].append({"nr": nr, "imie": w["imie"], "nazwisko": w["nazwisko"],
                                           "znizka": 0.0, "wplaty": [0.0] * len(self.dane["skladki"])})
            M.normalizuj(self.dane)
            self.zmieniono()

    def usun_ucznia(self):
        sel = self.tree_u.selection()
        if not sel or not sel[0].startswith("u"):
            messagebox.showinfo("Usuń", "Zaznacz ucznia na liście."); return
        u = self.dane["uczniowie"][int(sel[0][1:])]
        if messagebox.askyesno("Usuń ucznia", f'Usunąć {u["imie"]} {u["nazwisko"]} wraz z wpłatami?'):
            del self.dane["uczniowie"][int(sel[0][1:])]
            self.zmieniono()

    def dodaj_wplate(self):
        if not self.dane["uczniowie"]:
            messagebox.showinfo("Wpłata", "Najpierw dodaj uczniów."); return
        lista = [f'{u["nr"]}. {u["nazwisko"]} {u["imie"]}' for u in self.dane["uczniowie"]]
        sk = [s["nazwa"] for s in self.dane["skladki"]]

        def waliduj(w):
            try:
                if M.parse_kwota(w["kwota"]) <= 0:
                    return "Kwota musi być większa od 0."
            except ValueError:
                return "Wpisz kwotę jako liczbę, np. 150 lub 150,50."

        w = pytaj(self, "Dodaj wpłatę", [("uczen", "Uczeń", "combo", lista), ("skladka", "Składka", "combo", sk),
                                         ("kwota", "Kwota (zł)", "entry", None)], None, waliduj)
        if w:
            u = self.dane["uczniowie"][lista.index(w["uczen"])]
            u["wplaty"][sk.index(w["skladka"])] = round(u["wplaty"][sk.index(w["skladka"])] + M.parse_kwota(w["kwota"]), 2)
            self.zmieniono()

    # ------------------------------------------------------------------ REJESTR
    def _zbuduj_rejestr(self):
        f = self.tab_r
        pasek = ttk.Frame(f, padding=4)
        pasek.pack(fill="x")
        ttk.Button(pasek, text="➕ Dodaj wydatek", command=self.dodaj_wydatek).pack(side="left")
        ttk.Button(pasek, text="✏ Edytuj", command=self.edytuj_wydatek).pack(side="left", padx=6)
        ttk.Button(pasek, text="🗑 Usuń", command=self.usun_wydatek).pack(side="left")
        ttk.Label(pasek, text="   Filtr kategorii:").pack(side="left")
        self.filtr = ttk.Combobox(pasek, state="readonly", width=26, values=["(wszystkie)"])
        self.filtr.set("(wszystkie)")
        self.filtr.pack(side="left", padx=4)
        self.filtr.bind("<<ComboboxSelected>>", lambda e: self.odswiez_rejestr())
        self.lbl_razem = ttk.Label(pasek, font=("Segoe UI", 10, "bold"))
        self.lbl_razem.pack(side="right", padx=10)
        ramka = ttk.Frame(f)
        ramka.pack(fill="both", expand=True)
        kol = ("data", "dowod", "opis", "kat", "kwota", "forma", "osoba", "uwagi")
        self.tree_r = ttk.Treeview(ramka, columns=kol, show="headings", selectmode="browse")
        for k, n, w, a in [("data", "Data", 90, "center"), ("dowod", "Nr dowodu / faktury", 150, "w"),
                           ("opis", "Opis wydatku", 330, "w"), ("kat", "Kategoria", 170, "w"),
                           ("kwota", "Kwota", 110, "e"), ("forma", "Forma płatności", 110, "w"),
                           ("osoba", "Osoba rozliczająca", 150, "w"), ("uwagi", "Uwagi", 200, "w")]:
            self.tree_r.heading(k, text=n); self.tree_r.column(k, width=w, anchor=a)
        v = ttk.Scrollbar(ramka, orient="vertical", command=self.tree_r.yview)
        self.tree_r.configure(yscrollcommand=v.set)
        v.pack(side="right", fill="y")
        self.tree_r.pack(fill="both", expand=True)
        self.tree_r.bind("<Double-1>", lambda e: self.edytuj_wydatek())
        self.tree_r.tag_configure("brak", background="#FFD966")

    def odswiez_rejestr(self):
        d, t = self.dane, self.tree_r
        self.filtr["values"] = ["(wszystkie)"] + d["kategorie_wydatkow"]
        if self.filtr.get() not in self.filtr["values"]:
            self.filtr.set("(wszystkie)")
        y = t.yview()[0]
        t.delete(*t.get_children())
        suma = 0.0
        for i, w in sorted(enumerate(d["wydatki"]), key=lambda p: p[1]["data"]):
            if self.filtr.get() not in ("(wszystkie)", w["kategoria"]):
                continue
            suma += w["kwota"]
            # żółte tło = kategoria/forma spoza słownika (np. po usunięciu ze słownika)
            brak = w["kategoria"] not in d["kategorie_wydatkow"] or w["forma"] not in d["formy_platnosci"]
            t.insert("", "end", iid=f"w{i}", tags=("brak",) if brak else (),
                     values=(M.fmt_data(w["data"]), w["dowod"], w["opis"], w["kategoria"], M.fmt_zl(w["kwota"]),
                             w["forma"], w["osoba"], w["uwagi"]))
        self.lbl_razem.config(text=f"Razem (widoczne): {M.fmt_zl(suma)}")
        t.yview_moveto(y)

    def _formularz_wydatku(self, start=None):
        d = self.dane
        pola = [("data", "Data", "entry", None), ("dowod", "Nr dowodu / faktury", "entry", None),
                ("opis", "Opis wydatku", "entry", None), ("kategoria", "Kategoria", "combo", d["kategorie_wydatkow"]),
                ("kwota", "Kwota (zł)", "entry", None), ("forma", "Forma płatności", "combo", d["formy_platnosci"]),
                ("osoba", "Osoba rozliczająca", "entry", None), ("uwagi", "Uwagi", "entry", None)]

        def waliduj(w):
            try:
                M.parse_data(w["data"])
            except ValueError as e:
                return str(e)
            try:
                if M.parse_kwota(w["kwota"]) <= 0:
                    return "Kwota wydatku musi być większa od 0."
            except ValueError:
                return "Wpisz kwotę jako liczbę, np. 185,50."
            if not w["opis"]:
                return "Wpisz opis wydatku."
            return None

        return pytaj(self, "Wydatek", pola, start, waliduj)

    def dodaj_wydatek(self):
        ostatnia_osoba = self.dane["wydatki"][-1]["osoba"] if self.dane["wydatki"] else self.dane["parametry"]["skarbnik"]
        w = self._formularz_wydatku({"data": datetime.date.today().strftime("%d.%m.%Y"), "osoba": ostatnia_osoba})
        if w:
            self.dane["wydatki"].append(self._z_formularza(w))
            self.zmieniono()

    def _z_formularza(self, w):
        return {"data": M.parse_data(w["data"]), "dowod": w["dowod"], "opis": w["opis"], "kategoria": w["kategoria"],
                "kwota": M.parse_kwota(w["kwota"]), "forma": w["forma"], "osoba": w["osoba"], "uwagi": w["uwagi"]}

    def _wybrany_wydatek(self):
        sel = self.tree_r.selection()
        if not sel:
            messagebox.showinfo("Wydatek", "Zaznacz wydatek na liście."); return None
        return int(sel[0][1:])

    def edytuj_wydatek(self):
        i = self._wybrany_wydatek()
        if i is None:
            return
        x = self.dane["wydatki"][i]
        w = self._formularz_wydatku({**x, "data": M.fmt_data(x["data"]), "kwota": M.fmt_kw(x["kwota"]).replace(" ", "")})
        if w:
            self.dane["wydatki"][i] = self._z_formularza(w)
            self.zmieniono()

    def usun_wydatek(self):
        i = self._wybrany_wydatek()
        if i is not None and messagebox.askyesno("Usuń", f'Usunąć wydatek: {self.dane["wydatki"][i]["opis"]}?'):
            del self.dane["wydatki"][i]
            self.zmieniono()

    # ------------------------------------------------------------------ SLOWNIKI
    def _zbuduj_slowniki(self):
        f = self.tab_s
        ttk.Label(f, foreground="#595959", wraplength=1100, text=(
            "Każdy element w osobnej linii. Składki: 'Nazwa;stawka' (np. Mikołajki;20). Zmiana nazwy kategorii lub formy "
            "płatności (przy tej samej liczbie linii) automatycznie poprawia już wpisane wydatki. Kliknij 'Zastosuj'.")
                  ).grid(row=0, column=0, columnspan=4, sticky="w", padx=12, pady=8)
        self.txt = {}
        for col, (k, n) in enumerate([("kat", "Kategorie wydatków"), ("formy", "Formy płatności"),
                                      ("skl", "Składki (Nazwa;stawka)")]):
            lf = ttk.LabelFrame(f, text=n)
            lf.grid(row=1, column=col, padx=10, pady=6, sticky="n")
            t = tk.Text(lf, width=34 if k != "formy" else 22, height=10, font=("Segoe UI", 10))
            t.pack(padx=6, pady=6)
            self.txt[k] = t
        lf = ttk.LabelFrame(f, text="Parametry klasy")
        lf.grid(row=1, column=3, padx=10, pady=6, sticky="n")
        self.par = {}
        for i, (k, n) in enumerate([("klasa", "Klasa / szkoła"), ("rok", "Rok szkolny"), ("skarbnik", "Skarbnik"),
                                    ("saldo_poczatkowe", "Saldo początkowe (zł)"),
                                    ("inne_przychody", "Inne przychody (zł)")]):
            ttk.Label(lf, text=n).grid(row=i, column=0, sticky="e", padx=6, pady=4)
            e = ttk.Entry(lf, width=24); e.grid(row=i, column=1, padx=6, pady=4)
            self.par[k] = e
        ttk.Button(f, text="✔ Zastosuj zmiany", command=self.zastosuj_slowniki).grid(row=2, column=0, padx=12, pady=10, sticky="w")
        self._wczytaj_slowniki_do_pol()

    def _wczytaj_slowniki_do_pol(self):
        d = self.dane
        for k, lista in (("kat", d["kategorie_wydatkow"]), ("formy", d["formy_platnosci"]),
                         ("skl", [f'{s["nazwa"]};{s["stawka"]:g}' for s in d["skladki"]])):
            self.txt[k].delete("1.0", "end"); self.txt[k].insert("1.0", "\n".join(lista))
        for k, e in self.par.items():
            e.delete(0, "end")
            v = d["parametry"][k]
            e.insert(0, M.fmt_kw(v).replace(" ", "") if isinstance(v, float) else str(v))

    def zastosuj_slowniki(self):
        d = self.dane
        linie = lambda k: [l.strip() for l in self.txt[k].get("1.0", "end").splitlines() if l.strip()]
        kat, formy, skl = linie("kat"), linie("formy"), linie("skl")
        if not kat or not formy or not skl:
            messagebox.showwarning("Słowniki", "Każda lista musi mieć co najmniej jeden element."); return
        try:
            nowe_skl = []
            for l in skl:
                nazwa, _, stawka = l.partition(";")
                nowe_skl.append({"nazwa": nazwa.strip(), "stawka": M.parse_kwota(stawka)})
            par = {"klasa": self.par["klasa"].get().strip(), "rok": self.par["rok"].get().strip(),
                   "skarbnik": self.par["skarbnik"].get().strip(),
                   "saldo_poczatkowe": M.parse_kwota(self.par["saldo_poczatkowe"].get()),
                   "inne_przychody": M.parse_kwota(self.par["inne_przychody"].get())}
        except ValueError:
            messagebox.showwarning("Słowniki", "Stawki i kwoty muszą być liczbami (np. Mikołajki;20)."); return
        # Zmiana nazwy "w miejscu" (ta sama liczba linii) – przenosimy ją do istniejących wydatków
        for stare, nowe_l, pole in ((d["kategorie_wydatkow"], kat, "kategoria"), (d["formy_platnosci"], formy, "forma")):
            if len(stare) == len(nowe_l):
                mapa = {a: b for a, b in zip(stare, nowe_l) if a != b}
                for w in d["wydatki"]:
                    w[pole] = mapa.get(w[pole], w[pole])
        d["kategorie_wydatkow"], d["formy_platnosci"], d["skladki"], d["parametry"] = kat, formy, nowe_skl, par
        M.normalizuj(d)
        self._kolumny_uczniow()
        self.zmieniono()
        messagebox.showinfo("Słowniki", "Zapisano.")


if __name__ == "__main__":
    App().mainloop()

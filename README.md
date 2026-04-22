# 🚀 Student Delivery Processor

[![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python\&logoColor=white)](#)
[![Status](https://img.shields.io/badge/status-stable-success?style=flat\&logo=github)](#)
[![License](https://img.shields.io/badge/license-GPL--3.0-blue.svg)](#license)
[![Platform](https://img.shields.io/badge/platform-cross--platform-success?style=flat)](#)

<div align="center">

## 📚 Automazione intelligente delle consegne studenti

Trasforma cartelle Moodle disordinate + Excel Esse3 → struttura pulita, PDF e report Excel in pochi secondi.

</div>

---

## ✨ Perché questo progetto
Gestire manualmente centinaia di consegne studenti è lento, soggetto a errori e frustrante.

Questo tool automatizza tutto il processo:

* 🔍 riconosce automaticamente gli studenti
* 📂 sistema le cartelle
* 📄 genera un unico PDF per studente
* 📊 crea un report Excel pronto

---

## ⚡ Demo (prima → dopo)

### Input (Moodle)

```
MARIO LUIGI ROSSI_513901_assignsubmission_file
RAFFAELE DURSI_513835_assignsubmission_file
```

### Output

```
ROSSI MARIO LUIGI/
  └── LONGOBARDI_GIUSEPPE_PIO.pdf

D'URSI RAFFAELE/
  └── D'URSI_RAFFAELE.pdf
```

📊 File finale:

```
studenti_compilati.xlsx
```

---

## 📥 Input richiesto

### 📊 Excel Esse3 degli studenti frequentanti per anno accademico

File ufficiali esportati dalla piattaforma universitaria **Esse3** contenenti:

* Cognome e Nome
* Matricola
* Anno Accademico
* ecc... 

### 📂 Cartelle Moodle

Formato tipico:

```
NOME COGNOME_ID_assignsubmission_file
```

---

## 🧠 Matching intelligente

Il sistema utilizza più strategie combinate:

* ✔ Match su matricola
* ✔ Match nome esatto
* ✔ Match indipendente dall’ordine
* ✔ Supporto nomi multipli
* ✔ Normalizzazione caratteri (D'URSI = DURSI)

---

## ⚙️ Utilizzo

### ▶️ Script

```
python studenti_compatto.py
```

Modalità sicura:

```
python studenti_compatto.py --keep-originals
```

### 📓 Notebook

Apri:

```
workflow_studenti_compatto.ipynb
```

---

## 📦 Installazione

```bash
git clone https://github.com/giozoc/student-delivery-processor.git
cd student-delivery-processor
pip install -r requirements.txt
```

---

## 🎯 A chi è utile

* 🎓 Docenti universitari
* 👨‍🏫 Tutor e assistenti

---

## 📜 License

Questo progetto è distribuito sotto licenza **GNU GPL v3.0**.
--

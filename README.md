# Sammanställning
Det här programmet är skrivet i Python3.9

Programmet är skapat för att sammanställa Mentalbeskrivning Valp.
Som indata tar programmet en excel enligt tidigare känt format (saknar referens)

Som utdata skapar programmet en ny excel-fil med alla tillgängliga fält från mentalbeskrivningsformuläret
sammanställt valp för valp på enskilda rader.

## Installation
Först och främst så behöver ni Python3.9

För att sedan köra programmet så behövs ytterliggare moduler installeras.
För att installare alla beroenden.

gå in i sammanstall-mappen sedan
```bash
cd sammanstall 
pip install -r requirements.txt
```

## Körning

I sammanstall mappen finns en fil som heter `main.py`
Lägg in alla excel-filer som beskriver valparna i mappen `input`

Kör sedan

```bash
cd sammanstall 
python3.9 main.py 
```

När programmet är klart, så bör det finnas en fil i mappen `output` med sammanställningen.



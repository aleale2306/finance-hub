from datetime import date
import unittest

from finance_engine import (
    ANALYSIS_BUCKETS,
    analysis_bucket,
    historical_snapshots,
    monthly_actual_by_bucket,
    monthly_plan_by_bucket,
    uncategorized_transactions,
)


class FinanceEngineTests(unittest.TestCase):
    def test_analysis_buckets(self):
        cases = [
            ({"Importo": 100}, "Entrate"),
            ({"Importo": -1, "Tag": "Spese Fisse Inevitabili (NEEDS)"}, "Spese fisse inevitabili"),
            ({"Importo": -1, "NEED/WANT": "COMMITTED WANTS", "Fissa/Variabile": "Fissa"}, "Spese fisse evitabili"),
            ({"Importo": -1, "NEED/WANT": "NEEDS", "Fissa/Variabile": "Variabile"}, "Spese necessarie"),
            ({"Importo": -1, "NEED/WANT": "VARIABLE WANTS", "Fissa/Variabile": "Variabile"}, "Spese variabili non necessarie"),
            ({"Importo": -1, "Progetto": "Matrimonio"}, "Matrimonio"),
            ({"Importo": -1, "Categoria": "Shopping & Acquisti > Casa Nuova"}, "Casa nuova"),
            ({"Importo": -1, "Progetto": "Giappone", "Macro-categoria": "Viaggi"}, "Luna di miele (Giappone)"),
            ({"Importo": -1, "Macro-categoria": "Viaggi", "Progetto": "Sicilia"}, "Altri viaggi"),
        ]
        for row, expected in cases:
            with self.subTest(expected=expected):
                self.assertEqual(analysis_bucket(row), expected)

    def test_uncategorized_excludes_transfers(self):
        rows = [
            {"Data": date(2026, 1, 1), "Categoria": "Non classificata", "Trasferimento": False},
            {"Data": date(2026, 1, 2), "Categoria": "Non classificata", "Trasferimento": True},
            {"Data": date(2026, 1, 3), "Categoria": "Spesa", "Trasferimento": False},
        ]
        self.assertEqual(len(uncategorized_transactions(rows)), 1)

    def test_actual_monthly_aggregation(self):
        rows = [
            {"Data": date(2026, 1, 3), "Importo": 1000, "Trasferimento": False},
            {"Data": date(2026, 1, 4), "Importo": -200, "Trasferimento": False, "NEED/WANT": "NEEDS"},
            {"Data": date(2026, 1, 5), "Importo": -50, "Trasferimento": True, "NEED/WANT": "NEEDS"},
        ]
        result = monthly_actual_by_bucket(rows, 2026)
        self.assertEqual(set(result), set(ANALYSIS_BUCKETS))
        self.assertEqual(result["Entrate"][0], 1000)
        self.assertEqual(result["Spese necessarie"][0], 200)

    def test_monthly_plan(self):
        config = {
            "Entrate": [{"Importo": 1000, "Frequenza": "Mensile", "Stato": "Attiva"}],
            "Budget": [{"Anno": 2026, "Categoria MoneyWiz": "Spesa", "Importo budget (€)": 100, "Tipo budget": "Mensile", "Stato": "Confermato", "Classe di spesa": "Ricorrente variabile"}],
            "Mappatura Categorie": [{"Categoria MoneyWiz": "Spesa", "NEED/WANT": "NEEDS", "Fissa/Variabile": "Variabile", "Classe di spesa": "Ricorrente variabile"}],
        }
        result = monthly_plan_by_bucket(config, 2026)
        self.assertEqual(result["Entrate"], [1000] * 12)
        self.assertEqual(result["Spese necessarie"], [100] * 12)

    def test_historical_snapshots_include_partial_end(self):
        result = historical_snapshots(date(2026, 1, 10), date(2026, 3, 5))
        self.assertEqual(result, [date(2026, 1, 31), date(2026, 2, 28), date(2026, 3, 5)])


if __name__ == "__main__":
    unittest.main()

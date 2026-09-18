import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import mock

import pandas as pd

import laczenie_zamowien_formatki as app


class LaczenieFormatekTests(unittest.TestCase):
    def setUp(self):
        self.settings = {
            "TRYB_PRZYCINKI": app.TRYB_PRZYCINKI,
            "MAX_PRZYCINKA_PROC": app.MAX_PRZYCINKA_PROC,
            "MINIMUM_WARUNEK": app.MINIMUM_WARUNEK,
            "MIN_SZTUKI": app.MIN_SZTUKI,
            "MIN_M2": app.MIN_M2,
            "OPTYMALIZUJ_ODPAD": app.OPTYMALIZUJ_ODPAD,
            "MIN_OSZCZEDNOSC_ODPADU_M2": app.MIN_OSZCZEDNOSC_ODPADU_M2,
            "LACZ_DLUGOSC_Z_SZEROKOSCIA": app.LACZ_DLUGOSC_Z_SZEROKOSCIA,
            "LACZ_ROZNE_FEFCO": app.LACZ_ROZNE_FEFCO,
        }
        app.TRYB_PRZYCINKI = "side"
        app.MAX_PRZYCINKA_PROC = 10
        app.MINIMUM_WARUNEK = "OR"
        app.MIN_SZTUKI = 90
        app.MIN_M2 = 290
        app.OPTYMALIZUJ_ODPAD = True
        app.MIN_OSZCZEDNOSC_ODPADU_M2 = 0.01
        app.LACZ_DLUGOSC_Z_SZEROKOSCIA = False
        app.LACZ_ROZNE_FEFCO = False

    def tearDown(self):
        for name, value in self.settings.items():
            setattr(app, name, value)

    def test_wspolny_format_moze_laczyc_wymiary_po_obroceniu(self):
        df = pd.DataFrame(
            {
                "_effective_x": [100.0, 98.0],
                "_effective_y": [90.0, 92.0],
            }
        )

        target = app.best_target_for_indices(df, [0, 1])

        self.assertTrue(target.valid)
        self.assertEqual((target.target_x, target.target_y), (100.0, 92.0))
        self.assertLess(target.max_trim_pct, 2.18)

    def test_opcja_obrotu_laczy_dlugosc_z_szerokoscia(self):
        df = pd.DataFrame(
            {
                "_effective_x": [100.0, 52.0],
                "_effective_y": [50.0, 98.0],
            }
        )

        without_rotation = app.best_target_for_indices(df, [0, 1])
        app.LACZ_DLUGOSC_Z_SZEROKOSCIA = True
        with_rotation = app.best_target_for_indices(df, [0, 1])

        self.assertFalse(without_rotation.valid)
        self.assertTrue(with_rotation.valid)
        self.assertEqual((with_rotation.target_x, with_rotation.target_y), (100.0, 52.0))

    def test_opcja_roznych_fefco_tworzy_wspolna_grupe(self):
        source = pd.DataFrame(
            {
                "order_id": ["A", "B"],
                "position": [1, 1],
                "fefco": [201, 202],
                "dimension_x": [100, 100],
                "dimension_y": [50, 50],
                "quantity": [50, 50],
                "delivery_date": ["2026-07-30", "2026-07-31"],
            }
        )
        prepared, _, _ = app.przygotuj_dane(source, [])

        groups_separate, _, _ = app.znajdz_grupy_greedy(prepared)
        app.LACZ_ROZNE_FEFCO = True
        groups_mixed, assigned, _ = app.znajdz_grupy_greedy(prepared)

        self.assertTrue(groups_separate.empty)
        self.assertEqual(len(groups_mixed), 1)
        self.assertEqual(groups_mixed.loc[0, "fefco"], "201 + 202")
        self.assertEqual(len(assigned), 2)

    def test_alternative_dimension_tworzy_dwie_niezalezne_formatki(self):
        source = pd.DataFrame(
            {
                "order_id": ["ALT"],
                "position": [1],
                "fefco": [201],
                "dimension_x": [300],
                "dimension_y": [200],
                "quantity": [10],
                "delivery_date": ["2026-07-30"],
                "has_alternative_dimensions": [True],
                "alternative_dimension_x": [100],
                "alternative_dimension_y": [50],
            }
        )

        prepared, rejected, _ = app.przygotuj_dane(source, [])

        self.assertTrue(rejected.empty)
        self.assertEqual(prepared["_formatka_nr"].tolist(), [1, 2])
        self.assertEqual(prepared["_formatka_count"].tolist(), [2, 2])
        self.assertEqual(prepared["_formatka_count_for_type"].tolist(), [2, 2])
        self.assertEqual(prepared["_effective_quantity"].tolist(), [10.0, 10.0])
        self.assertEqual(prepared["_effective_quantity"].sum(), 20.0)

    def test_dwa_alternative_dimension_tworza_dwa_rodzaje_po_dwie_formatki(self):
        source = pd.DataFrame(
            {
                "order_id": ["ALT-2"],
                "position": [1],
                "fefco": [201],
                "dimension_x": [300],
                "dimension_y": [200],
                "quantity": [10],
                "delivery_date": ["2026-07-30"],
                "has_alternative_dimensions": [True],
                "alternative_dimension_x": [100],
                "alternative_dimension_y": [50],
                "has_alternative_dimensions_2": [True],
                "alternative_dimension_2_x": [80],
                "alternative_dimension_2_y": [40],
            }
        )

        prepared, rejected, _ = app.przygotuj_dane(source, [])

        self.assertTrue(rejected.empty)
        self.assertEqual(len(prepared), 4)
        self.assertEqual(
            prepared["_format_source"].tolist(),
            ["alternative_1", "alternative_1", "alternative_2", "alternative_2"],
        )
        self.assertEqual(prepared["_formatka_nr"].tolist(), [1, 2, 1, 2])
        self.assertEqual(prepared["_formatka_count_for_type"].tolist(), [2, 2, 2, 2])
        self.assertEqual(prepared["_formatka_count"].tolist(), [4, 4, 4, 4])
        self.assertEqual(
            prepared[["_effective_x", "_effective_y"]].values.tolist(),
            [[100.0, 50.0], [100.0, 50.0], [80.0, 40.0], [80.0, 40.0]],
        )
        self.assertEqual(prepared["_effective_quantity"].tolist(), [10.0] * 4)
        self.assertEqual(prepared["_effective_quantity"].sum(), 40.0)

    def test_dwie_formatki_moga_zostac_przypisane_do_grupy(self):
        source = pd.DataFrame(
            {
                "order_id": ["ALT", "ROT"],
                "position": [1, 1],
                "fefco": [201, 201],
                "dimension_x": [300, 98],
                "dimension_y": [200, 92],
                "quantity": [50, 20],
                "delivery_date": ["2026-07-30", "2026-08-01"],
                "has_alternative_dimensions": [True, False],
                "alternative_dimension_x": [100, None],
                "alternative_dimension_y": [90, None],
            }
        )
        prepared, _, _ = app.przygotuj_dane(source, [])

        groups, assigned, targets = app.znajdz_grupy_greedy(prepared)
        positions = app.przygotuj_pozycje_output(prepared, assigned, targets)

        self.assertEqual(len(assigned), 3)
        self.assertEqual(groups.loc[0, "liczba_pozycji"], 2)
        self.assertEqual(groups.loc[0, "liczba_formatek"], 3)
        self.assertEqual(groups.loc[0, "Sztuki do zamówienia"], 120.0)
        self.assertEqual(
            positions.loc[
                positions["Order ID"] == "ALT",
                "Numer formatki",
            ].tolist(),
            [1, 2],
        )
        self.assertEqual(set(positions["Status"]), {"POŁĄCZONO"})

    def test_obrocona_formatka_ma_flage_i_wyliczona_przycinke(self):
        app.LACZ_DLUGOSC_Z_SZEROKOSCIA = True
        source = pd.DataFrame(
            {
                "order_id": ["A", "B"],
                "position": [1, 1],
                "fefco": [202, 202],
                "dimension_x": [100, 52],
                "dimension_y": [50, 98],
                "quantity": [50, 50],
                "delivery_date": ["2026-07-30", "2026-07-31"],
            }
        )
        prepared, _, _ = app.przygotuj_dane(source, [])
        _, assigned, targets = app.znajdz_grupy_greedy(prepared)
        positions = app.przygotuj_pozycje_output(prepared, assigned, targets)

        rotated = positions.loc[positions["Order ID"] == "B"].iloc[0]
        self.assertEqual(rotated["Obrócono o 90°"], "TAK")
        self.assertNotEqual(rotated["Przycinka [%]"], "")

    def test_odrzucony_wiersz_ma_konkretny_powod(self):
        source = pd.DataFrame(
            {
                "order_id": ["BŁĄD"],
                "position": [1],
                "fefco": [""],
                "dimension_x": [0],
                "dimension_y": [50],
                "quantity": [0],
                "delivery_date": [None],
            }
        )
        _, rejected, _ = app.przygotuj_dane(source, [])
        output = app.przygotuj_odrzucone_output(rejected)

        reason = output.loc[0, "Powód odrzucenia"]
        self.assertIn("data dostawy", reason)
        self.assertIn("FEFCO", reason)
        self.assertIn("Długość", reason)
        self.assertIn("Ilość", reason)

    def test_optymalizacja_dzieli_grupe_gdy_obniza_laczny_odpad(self):
        df = pd.DataFrame(
            {
                "_effective_x": [190.0, 190.0, 200.0, 200.0],
                "_effective_y": [200.0, 200.0, 200.0, 200.0],
                "_effective_quantity": [50.0, 50.0, 50.0, 50.0],
                "_effective_m2_total": [1.9, 1.9, 2.0, 2.0],
                "_date": pd.to_datetime(["2026-07-30"] * 4),
            }
        )

        original_target = app.best_target_for_indices(df, [0, 1, 2, 3])
        original_waste = app.waste_metrics_for_indices(
            df,
            [0, 1, 2, 3],
            original_target,
        )["odpad_m2_razem"]

        optimized = app.optymalizuj_grupe_pod_katem_odpadu(df, [0, 1, 2, 3])
        optimized_waste = sum(
            app.waste_metrics_for_indices(
                df,
                group,
                app.best_target_for_indices(df, group),
            )["odpad_m2_razem"]
            for group in optimized
        )

        self.assertEqual([sorted(group) for group in optimized], [[0, 1], [2, 3]])
        self.assertLess(optimized_waste, original_waste)

    def test_optymalizacja_nie_dzieli_gdy_czesci_nie_spelniaja_minimum(self):
        df = pd.DataFrame(
            {
                "_effective_x": [190.0, 190.0, 200.0, 200.0],
                "_effective_y": [200.0, 200.0, 200.0, 200.0],
                "_effective_quantity": [30.0, 30.0, 30.0, 30.0],
                "_effective_m2_total": [1.14, 1.14, 1.2, 1.2],
                "_date": pd.to_datetime(["2026-07-30"] * 4),
            }
        )

        optimized = app.optymalizuj_grupe_pod_katem_odpadu(df, [0, 1, 2, 3])

        self.assertEqual(optimized, [[0, 1, 2, 3]])

    def test_optymalizacja_wydziela_samodzielna_formatke_spelniajaca_minima(self):
        app.MINIMUM_WARUNEK = "AND"
        df = pd.DataFrame(
            {
                # Przypadek odpowiadający zamówieniu 89276, pozycji 23.
                "_effective_x": [2566.0, 2566.0, 2566.0, 2566.0],
                "_effective_y": [1554.0, 1654.0, 1494.0, 1624.0],
                "_effective_quantity": [60.0, 60.0, 160.0, 20.0],
                "_effective_m2_total": [239.25384, 254.64984, 613.37664, 83.34368],
                "_date": pd.to_datetime(["2026-08-19"] * 4),
            }
        )

        optimized = app.optymalizuj_grupe_pod_katem_odpadu(df, [0, 1, 2, 3])

        self.assertIn([2], [sorted(group) for group in optimized])
        self.assertIn([0, 1, 3], [sorted(group) for group in optimized])

        original_target = app.best_target_for_indices(df, [0, 1, 2, 3])
        original_waste = app.waste_metrics_for_indices(
            df,
            [0, 1, 2, 3],
            original_target,
        )["odpad_m2_razem"]
        optimized_waste = sum(
            app.waste_metrics_for_indices(
                df,
                group,
                app.best_target_for_indices(df, group),
            )["odpad_m2_razem"]
            for group in optimized
        )

        self.assertLess(optimized_waste, original_waste)

    def test_podsumowanie_zawiera_formule_sumujaca_wszystkie_formatki(self):
        positions = pd.DataFrame(
            {
                "Status": ["POŁĄCZONO", "NIEPOŁĄCZONE"],
                "Powierzchnia wykorzystana [m²]": [10.0, 5.0],
                "Powierzchnia zamówiona łącznie [m²]": [11.0, ""],
                "Odpad łącznie [m²]": [1.0, ""],
            }
        )

        summary = app.przygotuj_podsumowanie(positions, pd.DataFrame())

        self.assertEqual(summary.loc[4, "Miara"], "SUMA M² WSZYSTKICH FORMATEK")
        self.assertEqual(summary.loc[4, "Wartość"], "=SUM('POZYCJE'!$B$2:$B$3)")
        self.assertIn("SUMIF", summary.loc[5, "Wartość"])

    def test_zapisuje_pod_nowa_nazwa_gdy_poprzedni_wynik_jest_zablokowany(self):
        with TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "zamowienia.xlsx"
            input_path.touch()
            expected_output = Path(temp_dir) / "zamowienia_polaczone_formatki.xlsx"
            expected_output.write_bytes(b"poprzedni wynik")

            empty = pd.DataFrame()
            positions = pd.DataFrame({"status": ["POŁĄCZONO"]})
            real_replace = Path.replace

            def replace_with_locked_target(source, target):
                if Path(target) == expected_output:
                    raise PermissionError("plik jest otwarty")
                return real_replace(source, target)

            with mock.patch.object(Path, "replace", autospec=True, side_effect=replace_with_locked_target):
                saved_path = app.zapisz_wynik(
                    input_path,
                    empty,
                    empty,
                    positions,
                    empty,
                    [],
                    "delivery_date",
                )

            self.assertNotEqual(saved_path, expected_output)
            self.assertTrue(saved_path.exists())
            self.assertTrue(saved_path.name.startswith("zamowienia_polaczone_formatki_"))
            self.assertEqual(expected_output.read_bytes(), b"poprzedni wynik")


if __name__ == "__main__":
    unittest.main()

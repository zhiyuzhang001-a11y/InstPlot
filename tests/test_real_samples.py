from pathlib import Path

import pytest

from instplot_io import read_data_file


FIXTURES = Path(__file__).parent / "fixtures" / "real_samples"


@pytest.mark.parametrize(
    ("name", "rows", "columns", "first_column", "last_column"),
    [
        ("0V_IP_1.dat", 405, 2, "B (Oe)", "M (emu)"),
        (
            "20240704_4_(2,-2)_20um_100uA_AHE(0V)L.dat",
            392,
            7,
            "Time (sec)",
            "Res_22_(Ohm)",
        ),
        ("3mA_0V_20um_6800Oe.txt", 184, 12, "时间", "Theta"),
        ("CoGd.txt", 89, 2, "Field(mT)", "GrayLevel"),
        ("Harm_3mA_4200Oe.txt", 374, 12, "时间", "Theta"),
        ("1V_2V_-2V.csv", 1795, 7, "index", "R_Ohm"),
    ],
)
def test_real_sample_columns_stay_aligned(name, rows, columns, first_column, last_column):
    result = read_data_file(FIXTURES / name)

    assert result.frame.shape == (rows, columns)
    assert result.frame.columns[0] == first_column
    assert result.frame.columns[-1] == last_column

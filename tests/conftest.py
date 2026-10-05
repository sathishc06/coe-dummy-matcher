import os
import sys
import zipfile

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
sys.path.insert(0, os.path.dirname(__file__))
import make_test_dataset as mk  # noqa: E402


@pytest.fixture(scope="session")
def dataset(tmp_path_factory):
    d = tmp_path_factory.mktemp("ds")
    return mk.build(str(d)) + (str(d),)


@pytest.fixture(scope="session")
def run(dataset, tmp_path_factory):
    from coe_matcher.pipeline import run_reconciliation
    vz, dz, idx, exp, _ = dataset
    return run_reconciliation(vz, dz, idx, str(tmp_path_factory.mktemp("w")), processes=1)


def subset_zip(src_zip, dst_zip, keep):
    with zipfile.ZipFile(src_zip) as zi, zipfile.ZipFile(dst_zip, "w") as zo:
        for n in zi.namelist():
            if any(k in n for k in keep):
                zo.writestr(n, zi.read(n))
    return dst_zip

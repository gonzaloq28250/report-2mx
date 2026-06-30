import shutil
import tempfile
import os
from openpyxl import load_workbook


def safe_save(wb, output_path):
    output_dir = os.path.dirname(output_path)
    os.makedirs(output_dir, exist_ok=True)

    fd, tmp_path = tempfile.mkstemp(suffix='.xlsx', dir=output_dir)
    os.close(fd)

    try:
        wb.save(tmp_path)
        wb.close()
        shutil.move(tmp_path, output_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def load_template(template_path, **kwargs):
    return load_workbook(template_path, **kwargs)

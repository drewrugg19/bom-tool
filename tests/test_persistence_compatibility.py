import importlib
import sqlite3
import subprocess
import sys
import tempfile
import threading
import unittest
from concurrent.futures import ThreadPoolExecutor
from contextlib import closing
from pathlib import Path
from unittest import mock

import pandas as pd
from FabBOMTool.core.export_files import reserve_export
from FabBOMTool.core import history, spool_reader, spool_overlay, logic


class PersistenceCompatibilityTests(unittest.TestCase):
    def test_parallel_export_reservations_preserve_existing_files(self):
        with tempfile.TemporaryDirectory() as directory:
            original = Path(directory) / 'shared.xlsx'
            original.write_bytes(b'original workbook')
            barrier = threading.Barrier(8)
            def export(i):
                barrier.wait()
                with reserve_export(directory, 'shared.xlsx') as path:
                    path.write_bytes(str(i).encode())
                    return path.name
            with ThreadPoolExecutor(max_workers=8) as pool:
                names = list(pool.map(export, range(8)))
            self.assertEqual(len(set(names)), 8)
            self.assertEqual(set(names), {f'shared_{i}.xlsx' for i in range(2,10)})
            self.assertEqual(original.read_bytes(), b'original workbook')
            self.assertEqual({(Path(directory)/name).read_bytes() for name in names}, {str(i).encode() for i in range(8)})

    def test_failed_export_removes_only_its_reservation(self):
        with tempfile.TemporaryDirectory() as directory:
            original = Path(directory)/'shared.xlsx'; original.write_bytes(b'original')
            with self.assertRaises(ValueError):
                with reserve_export(directory,'shared') as path:
                    path.write_bytes(b'partial'); raise ValueError('failed writer')
            self.assertEqual(original.read_bytes(),b'original')
            self.assertEqual(list(Path(directory).iterdir()),[original])

    def test_all_workflows_share_collision_safe_namespace_and_history(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory); original=root/'shared.xlsx'; original.write_bytes(b'old BOM export')
            with mock.patch.object(logic,'EXPORTS_DIR',root), mock.patch.object(logic,'load_legend_maps',return_value=({}, {}, {}, {}, set())), mock.patch.object(logic,'extract_from_pdf',return_value=pd.DataFrame()), mock.patch.object(history,'DB_PATH',root/'history.db'):
                original_id=history.save_run({'output_filename':'shared.xlsx'},[],'Company','')
                bom=logic.run_bom([],logic.DEFAULT_SETTINGS,'shared','Company','')
                row={column:'1' for column in spool_reader.SPOOL_COLUMNS}
                with mock.patch.object(spool_reader,'extract_spool_pdf',return_value=([row],[])):
                    standard=spool_reader.run_spool_reader(['standard.pdf'],'shared',root)
                with mock.patch.object(spool_overlay,'extract_spool_pdf',return_value=([row],[])):
                    overlay=spool_overlay.run_overlay(['old.pdf'],['new.pdf'],'shared',root)
                self.assertEqual([bom['output_filename'],standard['output_filename'],overlay['output_filename']],['shared_2.xlsx','shared_3.xlsx','shared_4.xlsx'])
                snapshots={result['output_filename']:Path(result['output']).read_bytes() for result in (bom,standard,overlay)}
                for result in (bom,standard,overlay):
                    run_id=history.save_run(result,[], 'Company','')
                    self.assertEqual(history.get_run(run_id)['export_file'],result['output_filename'])
                self.assertEqual(history.get_run(original_id)['export_file'],'shared.xlsx')
                self.assertEqual(original.read_bytes(),b'old BOM export')
                self.assertTrue(all((root/name).read_bytes()==data for name,data in snapshots.items()))

    def test_multiple_processes_migrate_legacy_database_without_row_changes(self):
        with tempfile.TemporaryDirectory() as directory:
            db=Path(directory)/'history.db'
            with closing(sqlite3.connect(db)) as con:
                con.execute("""CREATE TABLE runs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, run_date TEXT NOT NULL,
                    export_file TEXT NOT NULL, pdf_count INTEGER DEFAULT 0,
                    row_count INTEGER DEFAULT 0, total_inches REAL DEFAULT 0,
                    ok_rows INTEGER DEFAULT 0, warn_rows INTEGER DEFAULT 0,
                    err_rows INTEGER DEFAULT 0, mode TEXT DEFAULT 'Company',
                    project TEXT DEFAULT '', summary TEXT DEFAULT '', output_path TEXT DEFAULT '')""")
                con.execute("INSERT INTO runs (run_date, export_file, row_count, total_inches, project) VALUES ('original date','old.xlsx',7,12.5,'User project')")
                con.commit()
                columns=[r[1] for r in con.execute('PRAGMA table_info(runs)')]
                before=con.execute('SELECT * FROM runs').fetchall()
            script="""import importlib.util,sys
from pathlib import Path
spec=importlib.util.spec_from_file_location('migration',sys.argv[1])
m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
m.DB_PATH=Path(sys.argv[2])
for _ in range(10): m.init_db()
"""
            workers=[subprocess.Popen([sys.executable,'-c',script,history.__file__,str(db)],stdout=subprocess.PIPE,stderr=subprocess.PIPE) for _ in range(8)]
            for worker in workers:
                out,err=worker.communicate(timeout=45)
                self.assertEqual(worker.returncode,0,err.decode())
            with mock.patch.object(history,'DB_PATH',db):
                history.init_db(); history.init_db()
                self.assertEqual(history.get_run(1)['run_type'],'BOM / Inches')
            with closing(sqlite3.connect(db)) as con:
                self.assertEqual(con.execute('SELECT '+','.join(columns)+' FROM runs').fetchall(),before)
                names=[r[1] for r in con.execute('PRAGMA table_info(runs)')]
                self.assertEqual(set(names)-set(columns),{'run_type','pdf_filenames','run_metadata'})
                self.assertEqual(len(names),len(set(names)))

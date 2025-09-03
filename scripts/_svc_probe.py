import pathlib,datetime 
p=pathlib.Path(r"C:\ProgramData\FHRReportService\py_probe.txt") 
p.parent.mkdir(parents=True, exist_ok=True) 
p.write_text(f"OK at {datetime.datetime.now()}\n", encoding="utf-8") 

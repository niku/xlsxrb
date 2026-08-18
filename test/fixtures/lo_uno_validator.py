import sys
import os
import time
import subprocess
import argparse

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", required=True)
    parser.add_argument("--password", default="")
    args = parser.parse_args()

    port = 2002 + (os.getpid() % 1000)
    lo_proc = subprocess.Popen([
        "soffice",
        "--headless",
        f"--accept=socket,host=127.0.0.1,port={port};urp;",
        "-env:UserInstallation=file:///tmp/lo_uno_p_{}".format(os.getpid())
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    try:
        import uno
        from com.sun.star.beans import PropertyValue

        ctx = None
        for _ in range(30):
            try:
                local_ctx = uno.getComponentContext()
                smgr = local_ctx.ServiceManager
                resolver = smgr.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_ctx)
                ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.2)

        if not ctx:
            sys.stderr.write("Failed to connect to LibreOffice UNO bridge\n")
            sys.exit(1)

        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        file_url = uno.systemPathToFileUrl(os.path.abspath(args.file))
        props = []
        if args.password:
            p = PropertyValue()
            p.Name = "Password"
            p.Value = args.password
            props.append(p)

        hidden_prop = PropertyValue()
        hidden_prop.Name = "Hidden"
        hidden_prop.Value = True
        props.append(hidden_prop)

        try:
            doc = desktop.loadComponentFromURL(file_url, "_blank", 0, tuple(props))
        except Exception as e:
            sys.stderr.write(f"Load error: {e}\n")
            sys.exit(2)

        if not doc:
            sys.stderr.write("Document loaded as None\n")
            sys.exit(2)

        sheets = doc.getSheets()
        sheet = sheets.getByIndex(0)
        cell_a1 = sheet.getCellByPosition(0, 0).getString()
        cell_b1 = sheet.getCellByPosition(1, 0).getString()
        print(f"SUCCESS: Sheet 0, A1={cell_a1}, B1={cell_b1}")
        doc.close(True)
        sys.exit(0)
    finally:
        lo_proc.terminate()
        lo_proc.wait()

if __name__ == "__main__":
    main()

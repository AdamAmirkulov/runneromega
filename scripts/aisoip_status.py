import argparse
import time
import pathlib

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--workdir", required=True)
    p.add_argument("--iin", default="")
    p.add_argument("--login", default="")
    p.add_argument("--password", default="")
    args = p.parse_args()

    workdir = pathlib.Path(args.workdir)
    out = workdir / "out"
    out.mkdir(parents=True, exist_ok=True)

    print("AISOIP STATUS started")
    print("iin:", args.iin)
    print("login:", args.login)
    time.sleep(2)

    (out / "status_result.txt").write_text("demo status result\n", encoding="utf-8")
    print("AISOIP STATUS finished OK")

if __name__ == "__main__":
    main()

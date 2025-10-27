import sys, os, runpy

def set_cwd_to_bundle():
    # When frozen by PyInstaller, resources are under _MEIPASS; change CWD so relative paths work.
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    try:
        os.chdir(base)
    except Exception as e:
        # Fallback to current directory if something unexpected happens
        pass

def main():
    set_cwd_to_bundle()
    # Execute user's script as __main__ so if it has CLI/GUI startup on import, it still works.
    runpy.run_module("label_app", run_name="__main__")

if __name__ == "__main__":
    main()

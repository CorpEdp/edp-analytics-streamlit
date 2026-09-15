# modules/__init__.py
#
# This file is intentionally mostly empty.
# Its only job is to tell Python that "modules/" is an importable package,
# so that main.py can do:
#
#     import modules
#     pkgutil.iter_modules(modules.__path__)
#
# and auto-discover every .py file placed in this folder.
#
# You do NOT need to import each module here manually — main.py does that
# automatically for every file in this directory (except files starting
# with "_", like this one).

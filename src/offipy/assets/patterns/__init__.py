"""offipy.assets.patterns — deterministic procedural SVG pattern builders (A4).

Each pattern module exposes a small pure builder: given validated params plus
a seed, it returns an SVG template string (with theme sentinels) and the
template's color slot declarations. The provider in
``offipy.assets.providers.procedural`` owns schema validation and resolution.
"""

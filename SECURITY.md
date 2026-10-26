# Security policy

## Reporting a vulnerability

`zeroxlsx` writes files, it does not parse untrusted ones (there is no reader yet),
so the attack surface is small. Still, these are worth reporting:

- XML injection: a value or sheet name that breaks out of its element.
- Path handling: anything that makes `save()` or `add_image()` write outside the
  intended location.
- Memory exhaustion from a hostile input to `add_image()` (e.g. a header claiming an
  absurd pixel size).

Open a regular issue for anything harmless to disclose. For anything else, contact the
maintainer through the address on the repository profile and describe the pattern only.

## Scope

This is a file-format writer, not a security boundary. Do not feed it untrusted
binaries and expect it to validate them.

## Supported versions

The latest released minor version receives fixes.

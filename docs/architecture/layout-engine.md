# Layout Engine

`@pptkit/layout` should provide a controlled layout system for presentation generation.

It should not be framed as a full browser layout engine or a promise of complete CSS compatibility.

## Purpose

The layout layer exists so PPTKit can resolve placement and sizing as a first-class responsibility instead of burying layout inside authoring helpers or PPTX export code.

## First-Class Concepts

The layout package should be the home for concepts such as:

- stacking
- grid-like composition
- spacing and gaps
- alignment
- intrinsic sizing
- text measurement interfaces
- overflow detection
- pagination
- slide boundary checks

These are presentation-layout concerns, even when some of them resemble web concepts.

## Why It Deserves Its Own Package

Separating layout from export makes it easier to:

- test placement logic independently
- reuse layout behavior across preview and export paths
- avoid coupling file-format code to document composition
- evolve authoring ergonomics without rewriting exporters

## Controlled, Not Magical

The layout layer should remain understandable and predictable.

Contributors should be able to answer:

- what inputs the layout engine consumes
- what normalized placement output it produces
- which constraints it supports
- where unsupported cases fail or degrade

## Explicit Non-Goals

The initial architecture should not promise:

- full browser CSS parity
- arbitrary DOM fidelity
- support for every HTML and CSS behavior
- hidden magical auto-layout that cannot be reasoned about

PPTKit can learn from web layout models without committing to browser equivalence.

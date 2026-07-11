# Layout

Layout is a first-class concern in PPTKit.

## Why Layout Deserves Its Own Package

Legacy presentation libraries often push layout responsibility onto the user through coordinates and manual sizing.

PPTKit should treat layout as reusable logic rather than as a side effect of export.

## Responsibilities

The layout layer is expected to handle:

- Positioning
- Sizing
- Group composition
- Constraint resolution
- Placement output for renderers and exporters

## Design Goal

The layout package should provide consistent structure without becoming a hidden, overly magical system.

Contributors should be able to reason about:

- What inputs layout consumes
- What placement data it emits
- Where unsupported cases fail

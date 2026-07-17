# Plugins

PPTKit does not currently expose a stable plugin API.

Future extension work must define ownership, lifecycle, capability negotiation, version compatibility, and failure isolation before registration hooks become public. Plugins must not bypass Core validation or inject private OOXML concerns into the common authoring model.

Until that contract exists, compose higher-level helpers in application code using the public Core, Layout, and exporter APIs. Track plugin work in the [Roadmap](../../ROADMAP.md).

The external [PPTKit Presentation](https://github.com/openHacking/pptkit-presentation)
product is a higher-level workflow built only on public package APIs; it is not a
PPTKit plugin API or registration mechanism.

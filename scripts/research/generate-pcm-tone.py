#!/usr/bin/env python3
"""Generate the one-second 440 Hz mono PCM source for the native experiment."""
import math
import struct
import sys
import wave

with wave.open(sys.argv[1], "wb") as output:
    output.setnchannels(1)
    output.setsampwidth(2)
    output.setframerate(8000)
    output.writeframes(b"".join(struct.pack("<h", round(4000 * math.sin(2 * math.pi * 440 * index / 8000))) for index in range(8000)))

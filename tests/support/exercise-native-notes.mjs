/** Public-API contract exercised identically in Node and browser consumers. */
export async function exerciseNativeNotes(api, source, options) {
  const equal = (left, right, label) => {
    const stable = value => Array.isArray(value) ? value.map(stable) : value && typeof value === "object" ? Object.fromEntries(Object.keys(value).sort().map(key => [key, stable(value[key])])) : value;
    if (JSON.stringify(stable(left)) !== JSON.stringify(stable(right))) throw Error(label);
  };
  const bytes = Uint8Array.from(source), s = await api.importPptx(bytes, options);
  const checkNoOp = async () => {
    const plan = await s.prepareSave({ mode: "no-op" });
    if (plan.status !== "ready") throw Error("Native notes no-op blocked");
    equal([...(await s.save(plan.plan)).bytes], [...bytes], "Native notes no-op bytes changed");
  };
  try {
    const document = api.normalizePresentation(s.document), report = s.report;
    await checkNoOp();
    const first = s.document.slides[0], original = structuredClone(first.notes), rejections = [];
    for (const kind of ["font", "style", "language"]) {
      const wrong = structuredClone(original), run = wrong[0].runs.at(-1);
      run.text += "X";
      if (kind === "font") run.style.fontFamily = "Forbidden";
      if (kind === "style") run.style.bold = !run.style.bold;
      if (kind === "language") run.style.language = "zz-ZZ";
      first.setNotes(wrong);
      const blocked = await s.prepareSave({ mode: "edit" });
      if (blocked.status !== "blocked" || !blocked.report.diagnostics.length) throw Error("Mixed source notes " + kind + " write admitted");
      rejections.push({ kind, report: blocked.report });
      first.setNotes(original);
      await checkNoOp();
    }
    const candidates = [];
    for (let round = 1; round <= 3; round++) {
      for (const slide of s.document.slides) {
        const notes = structuredClone(slide.notes);
        notes[0].runs.at(-1).text += "X";
        slide.setNotes(notes);
      }
      const plan = await s.prepareSave({ mode: "edit" });
      if (plan.status !== "ready") throw Error("Native notes character candidate blocked");
      const saved = (await s.save(plan.plan)).bytes, candidate = await api.importPptx(saved, options);
      try {
        const document = api.normalizePresentation(candidate.document), report = candidate.report;
        equal(document, api.normalizePresentation(s.document), "Native notes full candidate projection differs");
        candidates.push({ round, bytes: [...saved], document, report });
      } finally { candidate.dispose(); }
    }
    return { document, report, exactNoOpAndRestores: true, rejections, candidates };
  } finally { s.dispose(); }
}

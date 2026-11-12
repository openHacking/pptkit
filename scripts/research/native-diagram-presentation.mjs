// Exact-layout research mapping, not a general DiagramML expression evaluator.
// Geometry remains the consuming application's responsibility.
export function diagramPresentationSpecimen(family, roots, nodes, ids, parentEdges, allocate, step) {
  const points = [], connections = [], childOrder = new Map(), sourceOrder = new Map(), destinationOrder = new Map();
  const byId = new Map(nodes.map(n => [n.id, n])), edgeByChild = new Map(parentEdges.map(e => [e.destination, e]));
  const point = (source, name, label, index = 0, count = 0) => {
    step(); const value = allocate();
    points.push({ id: value, source, name, ...(label === undefined ? {} : { label, index }), count }); return value;
  };
  const link = (type, source, destination) => {
    step();
    const sourceMap = type === "presParOf" ? childOrder : sourceOrder, order = sourceMap.get(source) ?? 0;
    const targetOrder = type === "presParOf" ? 0 : destinationOrder.get(destination) ?? 0;
    sourceMap.set(source, order + 1); if (type === "presOf") destinationOrder.set(destination, targetOrder + 1);
    connections.push({ id: allocate(), type, source, destination, order, targetOrder });
  };
  const child = (parent, source, name, label, index, count) => { const result = point(source, name, label, index, count); link("presParOf", parent, result); return result; };
  if (family === "process") {
    const flow = point(0, "linearFlow"); link("presOf", 0, flow);
    for (const [index, key] of roots.entries()) {
      step(); const source = ids.get(key), composite = child(flow, source, "composite");
      const title = child(composite, source, "parTx", "node1", 0, roots.length), shape = child(composite, source, "parSh", "node1", index, roots.length);
      const text = child(composite, source, "desTx", "fgAcc1", index, roots.length);
      link("presOf", source, title); link("presOf", source, shape);
      const stack = byId.get(key).children.slice().reverse();
      while (stack.length) {
        step(); const detail = byId.get(stack.pop()); link("presOf", ids.get(detail.id), text);
        for (let i = detail.children.length - 1; i >= 0; i--) stack.push(detail.children[i]);
      }
      if (index + 1 < roots.length) {
        const transition = edgeByChild.get(source).siblingTransition, view = child(flow, transition, "sibTrans", "sibTrans2D1", index, roots.length - 1);
        const connectorText = child(view, transition, "connTx", "sibTrans2D1", index, roots.length - 1);
        link("presOf", transition, view); link("presOf", transition, connectorText);
      }
    }
  } else if (family === "cycle") {
    const cycle = point(0, "cycle"); link("presOf", 0, cycle);
    for (const [index, key] of roots.entries()) {
      step(); const source = ids.get(key), view = child(cycle, source, "node", "node1", index, roots.length);
      link("presOf", source, view);
      const transition = edgeByChild.get(source).siblingTransition, connector = child(cycle, transition, "sibTrans", "sibTrans2D1", index, roots.length);
      const text = child(connector, transition, "connectorText", "sibTrans2D1", index, roots.length);
      link("presOf", transition, connector); link("presOf", transition, text);
    }
  } else if (family === "hierarchy") {
    const document = point(0, "hierChild1"); link("presOf", 0, document);
    const tasks = roots.map((key, index) => ({ key, index, count: roots.length, parent: document, depth: 1 })).reverse();
    while (tasks.length) {
      step(); const task = tasks.pop(), node = byId.get(task.key), source = ids.get(node.id), assistant = node.role === "assistant";
      if (task.transition !== undefined) {
        const transition = child(task.parent, task.transition, assistant ? "Name111" : "Name37", "parChTrans1D2", task.transitionIndex, task.transitionCount);
        link("presOf", task.transition, transition);
      }
      const suffix = assistant ? "3" : task.depth === 1 ? "1" : "", rootName = assistant ? "hierRoot3" : task.depth === 1 ? "hierRoot1" : "hierRoot2";
      const root = child(task.parent, source, rootName), composite = child(root, source, `rootComposite${suffix}`);
      const text = child(composite, source, `rootText${suffix}`, assistant ? "asst1" : task.depth === 1 ? "node0" : "node2", task.index, task.count);
      const connector = child(composite, source, `rootConnector${suffix}`, assistant ? "asst1" : task.depth === 1 ? "node1" : "node2", task.depth === 1 ? 0 : task.index, task.depth === 1 ? 0 : task.count);
      link("presOf", source, text); link("presOf", source, connector);
      const ordinary = child(root, source, assistant ? "hierChild6" : task.depth === 1 ? "hierChild2" : "hierChild4");
      const assistants = child(root, source, assistant ? "hierChild7" : task.depth === 1 ? "hierChild3" : "hierChild5");
      const normalKeys = node.children.filter(key => byId.get(key).role === "content"), assistantKeys = node.children.filter(key => byId.get(key).role === "assistant");
      const next = [];
      for (const [keys, parent] of [[normalKeys, ordinary], [assistantKeys, assistants]]) for (const [index, key] of keys.entries()) {
        step(); const edge = edgeByChild.get(ids.get(key));
        next.push({ key, index, count: keys.length, parent, depth: task.depth + 1, transition: edge.parentTransition, transitionIndex: edge.order, transitionCount: node.children.length });
      }
      for (let i = next.length - 1; i >= 0; i--) tasks.push(next[i]);
    }
  } else throw new TypeError("Unknown presentation family.");
  return { points, connections };
}

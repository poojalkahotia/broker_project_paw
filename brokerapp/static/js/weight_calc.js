function calcWeight({ partyId, millId, frkId, diffId, mode }) {
    const p = parseFloat(document.getElementById(partyId)?.value) || 0;
    const m = parseFloat(document.getElementById(millId)?.value) || 0;

    let frk = 0, diff = 0;

    if (p && m) {
        frk = (mode === 'sale') ? Math.max(p, m) : Math.min(p, m);
        diff = p - m;
    } else {
        frk = p || m;
        diff = 0;
    }

    document.getElementById(frkId).value = frk.toFixed(2);
    document.getElementById(diffId).value = diff.toFixed(2);
}

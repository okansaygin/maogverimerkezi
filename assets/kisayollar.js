// Veri Merkezi klavye kısayolları (Dash, assets/ klasöründeki .js dosyalarını otomatik yükler)
// Ctrl+K / Cmd+K : üst çubuktaki personel aramasına odaklan
document.addEventListener('keydown', function (e) {
  if ((e.ctrlKey || e.metaKey) && (e.key === 'k' || e.key === 'K')) {
    var kutu = document.querySelector('#global-arama input');
    if (kutu) { e.preventDefault(); kutu.focus(); }
  }
});

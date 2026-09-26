// Veri Merkezi - tarayıcı tanı kaydı
// Tarayıcıdaki çizim hataları (ör. bir bileşenin çökmesi) Dash'te normalde sessizdir; ama Dash böyle bir
// hatada sayfa durumunu geri alır ve kullanıcı "başka sayfaya atıldım" diye görür. Bu betik hataları ve
// tam sayfa yenilemelerini sunucudaki veri_merkezi_gunluk.log dosyasına yazdırır.
(function () {
  function gonder(tur, mesaj) {
    try {
      var veri = JSON.stringify({tur: tur, mesaj: String(mesaj).slice(0, 2000), yol: location.pathname});
      if (navigator.sendBeacon) {
        navigator.sendBeacon('/_tani', new Blob([veri], {type: 'application/json'}));
      } else {
        var x = new XMLHttpRequest(); x.open('POST', '/_tani', true);
        x.setRequestHeader('Content-Type', 'application/json'); x.send(veri);
      }
    } catch (e) { /* tanı kaydı uygulamayı asla bozmamalı */ }
  }
  window.addEventListener('error', function (e) {
    gonder('hata', (e.message || '') + (e.error && e.error.stack ? '\n' + e.error.stack : ''));
  });
  window.addEventListener('unhandledrejection', function (e) {
    gonder('promise', e.reason && (e.reason.stack || e.reason.message) || e.reason);
  });
  var asil = console.error;
  console.error = function () {
    try {
      var m = Array.prototype.map.call(arguments, function (a) { return a && a.stack ? a.stack : String(a); }).join(' ');
      if (!/Failed to load resource|ERR_|favicon/.test(m)) { gonder('console', m); }
    } catch (e) {}
    return asil.apply(console, arguments);
  };
  window.addEventListener('beforeunload', function () { gonder('yenileme', 'sayfa tamamen yenileniyor: ' + location.pathname); });
})();

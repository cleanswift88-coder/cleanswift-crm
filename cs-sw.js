// CleanSwift service worker: notifications locales + push (FCM)
self.addEventListener('install',()=>self.skipWaiting());
self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));
self.addEventListener('push',e=>{
  let raw={};
  try{raw=e.data?e.data.json():{};}catch(_){raw={title:'CleanSwift',body:e.data?e.data.text():''};}
  // FCM enveloppe le contenu dans data (ou notification)
  const d=(raw&&(raw.data||raw.notification))||raw||{};
  e.waitUntil(self.registration.showNotification(d.title||'CleanSwift',{
    body:d.body||'',icon:'images/icon-192.png',badge:'images/favicon-64.png',data:{url:d.url||'./'}
  }));
});
self.addEventListener('notificationclick',e=>{
  e.notification.close();
  const url=(e.notification.data&&e.notification.data.url)||'./';
  e.waitUntil(self.clients.matchAll({type:'window',includeUncontrolled:true}).then(cs=>{
    if(cs.length){return cs[0].focus();}
    return self.clients.openWindow(url);
  }));
});

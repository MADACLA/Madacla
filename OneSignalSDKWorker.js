// El "cartero" de los avisos de novedades, para madacla.es.
//
// Claudia (2026-09-29): «a la gente que se mete la web en la pantalla de
// inicio del movil, ¿hay forma de que le llegue un aviso cuando subimos
// novedades?». Si: avisos web (web push). En el iPhone SOLO funcionan si la
// clienta se ha puesto la web en la pantalla de inicio, que es justo lo que
// ella hace con sus clientas.
//
// Este fichero tiene que estar en la RAIZ del sitio (madacla.es/OneSignalSDKWorker.js)
// y con este nombre exacto: lo busca ahi el SDK. No lleva nada nuestro, solo
// carga el programa de OneSignal.
//
// Va en publicar.sh, como index.html y estilos.css. Si algun dia deja de
// copiarse, los avisos dejan de funcionar sin dar la cara.
importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js");

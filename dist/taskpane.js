!function(e){var t={};function n(r){if(t[r])return t[r].exports;var o=t[r]={i:r,l:!1,exports:{}};return e[r].call(o.exports,o,o.exports,n),o.l=!0,o.exports}n.m=e,n.c=t,n.d=function(e,t,r){n.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:r})},n.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},n.t=function(e,t){if(1&t&&(e=n(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var r=Object.create(null);if(n.r(r),Object.defineProperty(r,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var o in e)n.d(r,o,function(t){return e[t]}.bind(null,o));return r},n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,"a",t),t},n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},n.p="",n(n.s=307)}({307:function(e,t,n){"use strict";function r(e,t,n,r,o,u,i){try{var c=e[u](i),a=c.value}catch(e){return void n(e)}c.done?t(a):Promise.resolve(a).then(r,o)}n.r(t),n.d(t,"run",function(){return u});var o="https://dialog-opener.surge.sh/dialog.html";function u(){return i.apply(this,arguments)}function i(){var e;return e=regeneratorRuntime.mark(function e(){return regeneratorRuntime.wrap(function(e){for(;;)switch(e.prev=e.next){case 0:Office.context.ui.displayDialogAsync(o);case 1:case"end":return e.stop()}},e)}),(i=function(){var t=this,n=arguments;return new Promise(function(o,u){var i=e.apply(t,n);function c(e){r(i,o,u,c,a,"next",e)}function a(e){r(i,o,u,c,a,"throw",e)}c(void 0)})}).apply(this,arguments)}Office.onReady(function(e){e.host===Office.HostType.Outlook&&(document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("run").onclick=u)})}});
//# sourceMappingURL=taskpane.js.map
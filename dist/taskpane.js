/*! For license information please see taskpane.js.LICENSE.txt */
!function(){var t={84400:function(t,e,r){"use strict";t.exports=r.p+"assets/copilot.png"},60947:function(t,e,r){"use strict";t.exports=r.p+"d0f8cd0c30a67a7c0fae.js"},58394:function(t,e,r){"use strict";t.exports=r.p+"3437828449654ee32f4d.css"}},e={};function r(n){var o=e[n];if(void 0!==o)return o.exports;var i=e[n]={exports:{}};return t[n](i,i.exports,r),i.exports}r.m=t,r.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),r.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;r.g.importScripts&&(t=r.g.location+"");var e=r.g.document;if(!t&&e&&(e.currentScript&&"SCRIPT"===e.currentScript.tagName.toUpperCase()&&(t=e.currentScript.src),!t)){var n=e.getElementsByTagName("script");if(n.length)for(var o=n.length-1;o>-1&&(!t||!/^http(s?):/.test(t));)t=n[o--].src}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),r.p=t}(),r.b=document.baseURI||self.location.href,function(){function t(e){return t="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},t(e)}function e(){"use strict";e=function(){return n};var r,n={},o=Object.prototype,i=o.hasOwnProperty,c=Object.defineProperty||function(t,e,r){t[e]=r.value},a="function"==typeof Symbol?Symbol:{},u=a.iterator||"@@iterator",s=a.asyncIterator||"@@asyncIterator",f=a.toStringTag||"@@toStringTag";function l(t,e,r){return Object.defineProperty(t,e,{value:r,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{l({},"")}catch(r){l=function(t,e,r){return t[e]=r}}function p(t,e,r,n){var o=e&&e.prototype instanceof w?e:w,i=Object.create(o.prototype),a=new N(n||[]);return c(i,"_invoke",{value:T(t,r,a)}),i}function h(t,e,r){try{return{type:"normal",arg:t.call(e,r)}}catch(t){return{type:"throw",arg:t}}}n.wrap=p;var d="suspendedStart",v="suspendedYield",y="executing",m="completed",g={};function w(){}function b(){}function x(){}var E={};l(E,u,(function(){return this}));var L=Object.getPrototypeOf,k=L&&L(L(P([])));k&&k!==o&&i.call(k,u)&&(E=k);var I=x.prototype=w.prototype=Object.create(E);function O(t){["next","throw","return"].forEach((function(e){l(t,e,(function(t){return this._invoke(e,t)}))}))}function j(e,r){function n(o,c,a,u){var s=h(e[o],e,c);if("throw"!==s.type){var f=s.arg,l=f.value;return l&&"object"==t(l)&&i.call(l,"__await")?r.resolve(l.__await).then((function(t){n("next",t,a,u)}),(function(t){n("throw",t,a,u)})):r.resolve(l).then((function(t){f.value=t,a(f)}),(function(t){return n("throw",t,a,u)}))}u(s.arg)}var o;c(this,"_invoke",{value:function(t,e){function i(){return new r((function(r,o){n(t,e,r,o)}))}return o=o?o.then(i,i):i()}})}function T(t,e,n){var o=d;return function(i,c){if(o===y)throw Error("Generator is already running");if(o===m){if("throw"===i)throw c;return{value:r,done:!0}}for(n.method=i,n.arg=c;;){var a=n.delegate;if(a){var u=B(a,n);if(u){if(u===g)continue;return u}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if(o===d)throw o=m,n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);o=y;var s=h(t,e,n);if("normal"===s.type){if(o=n.done?m:v,s.arg===g)continue;return{value:s.arg,done:n.done}}"throw"===s.type&&(o=m,n.method="throw",n.arg=s.arg)}}}function B(t,e){var n=e.method,o=t.iterator[n];if(o===r)return e.delegate=null,"throw"===n&&t.iterator.return&&(e.method="return",e.arg=r,B(t,e),"throw"===e.method)||"return"!==n&&(e.method="throw",e.arg=new TypeError("The iterator does not provide a '"+n+"' method")),g;var i=h(o,t.iterator,e.arg);if("throw"===i.type)return e.method="throw",e.arg=i.arg,e.delegate=null,g;var c=i.arg;return c?c.done?(e[t.resultName]=c.value,e.next=t.nextLoc,"return"!==e.method&&(e.method="next",e.arg=r),e.delegate=null,g):c:(e.method="throw",e.arg=new TypeError("iterator result is not an object"),e.delegate=null,g)}function S(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function _(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function N(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(S,this),this.reset(!0)}function P(e){if(e||""===e){var n=e[u];if(n)return n.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var o=-1,c=function t(){for(;++o<e.length;)if(i.call(e,o))return t.value=e[o],t.done=!1,t;return t.value=r,t.done=!0,t};return c.next=c}}throw new TypeError(t(e)+" is not iterable")}return b.prototype=x,c(I,"constructor",{value:x,configurable:!0}),c(x,"constructor",{value:b,configurable:!0}),b.displayName=l(x,f,"GeneratorFunction"),n.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===b||"GeneratorFunction"===(e.displayName||e.name))},n.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,x):(t.__proto__=x,l(t,f,"GeneratorFunction")),t.prototype=Object.create(I),t},n.awrap=function(t){return{__await:t}},O(j.prototype),l(j.prototype,s,(function(){return this})),n.AsyncIterator=j,n.async=function(t,e,r,o,i){void 0===i&&(i=Promise);var c=new j(p(t,e,r,o),i);return n.isGeneratorFunction(e)?c:c.next().then((function(t){return t.done?t.value:c.next()}))},O(I),l(I,f,"Generator"),l(I,u,(function(){return this})),l(I,"toString",(function(){return"[object Generator]"})),n.keys=function(t){var e=Object(t),r=[];for(var n in e)r.push(n);return r.reverse(),function t(){for(;r.length;){var n=r.pop();if(n in e)return t.value=n,t.done=!1,t}return t.done=!0,t}},n.values=P,N.prototype={constructor:N,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=r,this.done=!1,this.delegate=null,this.method="next",this.arg=r,this.tryEntries.forEach(_),!t)for(var e in this)"t"===e.charAt(0)&&i.call(this,e)&&!isNaN(+e.slice(1))&&(this[e]=r)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var e=this;function n(n,o){return a.type="throw",a.arg=t,e.next=n,o&&(e.method="next",e.arg=r),!!o}for(var o=this.tryEntries.length-1;o>=0;--o){var c=this.tryEntries[o],a=c.completion;if("root"===c.tryLoc)return n("end");if(c.tryLoc<=this.prev){var u=i.call(c,"catchLoc"),s=i.call(c,"finallyLoc");if(u&&s){if(this.prev<c.catchLoc)return n(c.catchLoc,!0);if(this.prev<c.finallyLoc)return n(c.finallyLoc)}else if(u){if(this.prev<c.catchLoc)return n(c.catchLoc,!0)}else{if(!s)throw Error("try statement without catch or finally");if(this.prev<c.finallyLoc)return n(c.finallyLoc)}}}},abrupt:function(t,e){for(var r=this.tryEntries.length-1;r>=0;--r){var n=this.tryEntries[r];if(n.tryLoc<=this.prev&&i.call(n,"finallyLoc")&&this.prev<n.finallyLoc){var o=n;break}}o&&("break"===t||"continue"===t)&&o.tryLoc<=e&&e<=o.finallyLoc&&(o=null);var c=o?o.completion:{};return c.type=t,c.arg=e,o?(this.method="next",this.next=o.finallyLoc,g):this.complete(c)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),g},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.finallyLoc===t)return this.complete(r.completion,r.afterLoc),_(r),g}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.tryLoc===t){var n=r.completion;if("throw"===n.type){var o=n.arg;_(r)}return o}}throw Error("illegal catch attempt")},delegateYield:function(t,e,n){return this.delegate={iterator:P(t),resultName:e,nextLoc:n},"next"===this.method&&(this.arg=r),g}},n}function r(t,e,r,n,o,i,c){try{var a=t[i](c),u=a.value}catch(t){return void r(t)}a.done?e(u):Promise.resolve(u).then(n,o)}function n(t){return function(){var e=this,n=arguments;return new Promise((function(o,i){var c=t.apply(e,n);function a(t){r(c,o,i,a,u,"next",t)}function u(t){r(c,o,i,a,u,"throw",t)}a(void 0)}))}}function o(t,e,r){var n=document.getElementById("chatWindow");e&&e.attachments&&e.attachments.length>0?e.attachments.forEach((function(t){t.content&&t.content.buttons&&t.content.buttons.length>0&&t.content.buttons.forEach((function(e){if("signin"===e.type){var r=document.createElement("button");r.innerText=e.title||"Sign In",r.classList.add("ms-Button","ms-Button--primary"),r.onclick=function(){window.open(e.value,"_blank")},n.innerHTML+='<div class="bot"><img src="assets/copilot.png" alt="Copilot Icon" /> <br>'.concat(t.content.text,"</div>"),n.appendChild(r)}}))})):"bot"===r?e.text&&(n.innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">'.concat(e.text,"</div>")):t&&(n.innerHTML+='<div class="user-wrapper">You</div><div class="message user">'.concat(t,"</div>")),document.getElementById("userInput").value=""}function i(t){return c.apply(this,arguments)}function c(){return c=n(e().mark((function t(r){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=n(e().mark((function t(n){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n.document.body.insertText(r,Word.InsertLocation.end),t.next=4,n.sync();case 4:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:case"end":return t.stop()}}),t)}))),c.apply(this,arguments)}Office.onReady((function(t){if(t.host===Office.HostType.Word){var r=null,c=!0;document.getElementById("myButton").onclick=function(){Office.context.ui.displayDialogAsync('https://krishnamoorthykana.github.io/NoviWord/src/taskpane/index.html"',{height:30,width:20})},document.addEventListener("DOMContentLoaded",n(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if(!c){t.next=5;break}return t.next=3,a();case 3:r=t.sent,c=!1;case 5:case"end":return t.stop()}}),t)})))),document.getElementById("askButton").onclick=n(e().mark((function t(){var n;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if(!(n=document.getElementById("userInput").value)){t.next=5;break}return o(n,"","User"),t.next=5,u(r,n);case 5:case"end":return t.stop()}}),t)}))),document.getElementById("userInput").addEventListener("keydown",function(){var t=n(e().mark((function t(n){var i;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if("Enter"!==n.key){t.next=7;break}if(n.preventDefault(),!(i=document.getElementById("userInput").value)){t.next=7;break}return o(i,"","User"),t.next=7,u(r,i);case 7:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()),document.getElementById("insertButton").onclick=n(e().mark((function t(){var r;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if(!(r=document.getElementById("chatWindow").lastChild?document.getElementById("chatWindow").lastChild.innerText:"")){t.next=4;break}return t.next=4,i(r);case 4:case"end":return t.stop()}}),t)})))}}));var a=function(){var t=n(e().mark((function t(){var r,n,i;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,t.next=3,fetch("https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview");case 3:return r=t.sent,t.next=6,r.json();case 6:if(n=t.sent,(i=new window.DirectLine.DirectLine({token:n.token}))&&i.activity$){t.next=10;break}throw new Error("DirectLine instance failed to initialize");case 10:return i.postActivity({from:{id:"10",name:"User"},type:"message",text:"Hi"}).subscribe((function(t){return console.log("Message sent with ID:",t)}),(function(t){return console.error("Error sending message:",t)})),i.activity$.subscribe((function(t){console.log("Testing activity: ",t),console.log("Role",t.from.role),"message"!==t.type||"10"===t.from.id||t.recipient||(console.log("Bot Response: ",t.text),o(!1,t,t.from.role))})),t.abrupt("return",i);case 15:t.prev=15,t.t0=t.catch(0),console.error("Error initializing DirectLine:",t.t0);case 18:case"end":return t.stop()}}),t,null,[[0,15]])})));return function(){return t.apply(this,arguments)}}(),u=function(){var t=n(e().mark((function t(r,n){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:r.postActivity({from:{id:"10",name:"User"},type:"message",text:n}).subscribe((function(t){return console.log("Message sent with ID:",t)}),(function(t){return console.error("Error sending message:",t)}));case 1:case"end":return t.stop()}}),t)})));return function(e,r){return t.apply(this,arguments)}}()}(),function(){"use strict";new URL(r(58394),r.b),new URL(r(84400),r.b),new URL(r(60947),r.b)}()}();
//# sourceMappingURL=taskpane.js.map
/*! For license information please see taskpane.js.LICENSE.txt */
!function(){var t={84400:function(t,e,n){"use strict";t.exports=n.p+"assets/copilot.png"},60947:function(t,e,n){"use strict";t.exports=n.p+"d8adffda581263d61531.js"},58394:function(t,e,n){"use strict";t.exports=n.p+"faca41d22060e85c0a8c.css"}},e={};function n(r){var o=e[r];if(void 0!==o)return o.exports;var i=e[r]={exports:{}};return t[r](i,i.exports,n),i.exports}n.m=t,n.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),n.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;n.g.importScripts&&(t=n.g.location+"");var e=n.g.document;if(!t&&e&&(e.currentScript&&"SCRIPT"===e.currentScript.tagName.toUpperCase()&&(t=e.currentScript.src),!t)){var r=e.getElementsByTagName("script");if(r.length)for(var o=r.length-1;o>-1&&(!t||!/^http(s?):/.test(t));)t=r[o--].src}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),n.p=t}(),n.b=document.baseURI||self.location.href,function(){function t(e){return t="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},t(e)}function e(){"use strict";e=function(){return r};var n,r={},o=Object.prototype,i=o.hasOwnProperty,c=Object.defineProperty||function(t,e,n){t[e]=n.value},a="function"==typeof Symbol?Symbol:{},s=a.iterator||"@@iterator",u=a.asyncIterator||"@@asyncIterator",l=a.toStringTag||"@@toStringTag";function f(t,e,n){return Object.defineProperty(t,e,{value:n,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{f({},"")}catch(n){f=function(t,e,n){return t[e]=n}}function p(t,e,n,r){var o=e&&e.prototype instanceof w?e:w,i=Object.create(o.prototype),a=new N(r||[]);return c(i,"_invoke",{value:j(t,n,a)}),i}function h(t,e,n){try{return{type:"normal",arg:t.call(e,n)}}catch(t){return{type:"throw",arg:t}}}r.wrap=p;var d="suspendedStart",v="suspendedYield",y="executing",m="completed",g={};function w(){}function b(){}function x(){}var E={};f(E,s,(function(){return this}));var L=Object.getPrototypeOf,k=L&&L(L(P([])));k&&k!==o&&i.call(k,s)&&(E=k);var I=x.prototype=w.prototype=Object.create(E);function T(t){["next","throw","return"].forEach((function(e){f(t,e,(function(t){return this._invoke(e,t)}))}))}function B(e,n){function r(o,c,a,s){var u=h(e[o],e,c);if("throw"!==u.type){var l=u.arg,f=l.value;return f&&"object"==t(f)&&i.call(f,"__await")?n.resolve(f.__await).then((function(t){r("next",t,a,s)}),(function(t){r("throw",t,a,s)})):n.resolve(f).then((function(t){l.value=t,a(l)}),(function(t){return r("throw",t,a,s)}))}s(u.arg)}var o;c(this,"_invoke",{value:function(t,e){function i(){return new n((function(n,o){r(t,e,n,o)}))}return o=o?o.then(i,i):i()}})}function j(t,e,r){var o=d;return function(i,c){if(o===y)throw Error("Generator is already running");if(o===m){if("throw"===i)throw c;return{value:n,done:!0}}for(r.method=i,r.arg=c;;){var a=r.delegate;if(a){var s=O(a,r);if(s){if(s===g)continue;return s}}if("next"===r.method)r.sent=r._sent=r.arg;else if("throw"===r.method){if(o===d)throw o=m,r.arg;r.dispatchException(r.arg)}else"return"===r.method&&r.abrupt("return",r.arg);o=y;var u=h(t,e,r);if("normal"===u.type){if(o=r.done?m:v,u.arg===g)continue;return{value:u.arg,done:r.done}}"throw"===u.type&&(o=m,r.method="throw",r.arg=u.arg)}}}function O(t,e){var r=e.method,o=t.iterator[r];if(o===n)return e.delegate=null,"throw"===r&&t.iterator.return&&(e.method="return",e.arg=n,O(t,e),"throw"===e.method)||"return"!==r&&(e.method="throw",e.arg=new TypeError("The iterator does not provide a '"+r+"' method")),g;var i=h(o,t.iterator,e.arg);if("throw"===i.type)return e.method="throw",e.arg=i.arg,e.delegate=null,g;var c=i.arg;return c?c.done?(e[t.resultName]=c.value,e.next=t.nextLoc,"return"!==e.method&&(e.method="next",e.arg=n),e.delegate=null,g):c:(e.method="throw",e.arg=new TypeError("iterator result is not an object"),e.delegate=null,g)}function S(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function _(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function N(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(S,this),this.reset(!0)}function P(e){if(e||""===e){var r=e[s];if(r)return r.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var o=-1,c=function t(){for(;++o<e.length;)if(i.call(e,o))return t.value=e[o],t.done=!1,t;return t.value=n,t.done=!0,t};return c.next=c}}throw new TypeError(t(e)+" is not iterable")}return b.prototype=x,c(I,"constructor",{value:x,configurable:!0}),c(x,"constructor",{value:b,configurable:!0}),b.displayName=f(x,l,"GeneratorFunction"),r.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===b||"GeneratorFunction"===(e.displayName||e.name))},r.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,x):(t.__proto__=x,f(t,l,"GeneratorFunction")),t.prototype=Object.create(I),t},r.awrap=function(t){return{__await:t}},T(B.prototype),f(B.prototype,u,(function(){return this})),r.AsyncIterator=B,r.async=function(t,e,n,o,i){void 0===i&&(i=Promise);var c=new B(p(t,e,n,o),i);return r.isGeneratorFunction(e)?c:c.next().then((function(t){return t.done?t.value:c.next()}))},T(I),f(I,l,"Generator"),f(I,s,(function(){return this})),f(I,"toString",(function(){return"[object Generator]"})),r.keys=function(t){var e=Object(t),n=[];for(var r in e)n.push(r);return n.reverse(),function t(){for(;n.length;){var r=n.pop();if(r in e)return t.value=r,t.done=!1,t}return t.done=!0,t}},r.values=P,N.prototype={constructor:N,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=n,this.done=!1,this.delegate=null,this.method="next",this.arg=n,this.tryEntries.forEach(_),!t)for(var e in this)"t"===e.charAt(0)&&i.call(this,e)&&!isNaN(+e.slice(1))&&(this[e]=n)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var e=this;function r(r,o){return a.type="throw",a.arg=t,e.next=r,o&&(e.method="next",e.arg=n),!!o}for(var o=this.tryEntries.length-1;o>=0;--o){var c=this.tryEntries[o],a=c.completion;if("root"===c.tryLoc)return r("end");if(c.tryLoc<=this.prev){var s=i.call(c,"catchLoc"),u=i.call(c,"finallyLoc");if(s&&u){if(this.prev<c.catchLoc)return r(c.catchLoc,!0);if(this.prev<c.finallyLoc)return r(c.finallyLoc)}else if(s){if(this.prev<c.catchLoc)return r(c.catchLoc,!0)}else{if(!u)throw Error("try statement without catch or finally");if(this.prev<c.finallyLoc)return r(c.finallyLoc)}}}},abrupt:function(t,e){for(var n=this.tryEntries.length-1;n>=0;--n){var r=this.tryEntries[n];if(r.tryLoc<=this.prev&&i.call(r,"finallyLoc")&&this.prev<r.finallyLoc){var o=r;break}}o&&("break"===t||"continue"===t)&&o.tryLoc<=e&&e<=o.finallyLoc&&(o=null);var c=o?o.completion:{};return c.type=t,c.arg=e,o?(this.method="next",this.next=o.finallyLoc,g):this.complete(c)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),g},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.finallyLoc===t)return this.complete(n.completion,n.afterLoc),_(n),g}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.tryLoc===t){var r=n.completion;if("throw"===r.type){var o=r.arg;_(n)}return o}}throw Error("illegal catch attempt")},delegateYield:function(t,e,r){return this.delegate={iterator:P(t),resultName:e,nextLoc:r},"next"===this.method&&(this.arg=n),g}},r}function n(t,e,n,r,o,i,c){try{var a=t[i](c),s=a.value}catch(t){return void n(t)}a.done?e(s):Promise.resolve(s).then(r,o)}function r(t){return function(){var e=this,r=arguments;return new Promise((function(o,i){var c=t.apply(e,r);function a(t){n(c,o,i,a,s,"next",t)}function s(t){n(c,o,i,a,s,"throw",t)}a(void 0)}))}}function o(t,e,n){var r=document.getElementById("chatWindow");e&&e.attachments&&e.attachments.length>0?e.attachments.forEach((function(t){t.content&&t.content.buttons&&t.content.buttons.length>0&&t.content.buttons.forEach((function(e){if("signin"===e.type){var n=document.createElement("button");n.innerText=e.title||"Sign In",n.classList.add("ms-Button","ms-Button--primary"),n.onclick=function(){window.open(e.value,"_blank")},r.innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">'.concat(t.content.text,"</div>"),r.appendChild(n)}}))})):"bot"===n?e.text&&(r.innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">'.concat(e.text,"</div>")):t&&(r.innerHTML+='<div class="user-wrapper">You</div><div class="message user">'.concat(t,"</div>")),function(){var t=document.getElementById("chatWindow");setTimeout((function(){t.scrollTop=t.scrollHeight}),100)}(),document.getElementById("userInput").value=""}function i(t){return c.apply(this,arguments)}function c(){return c=r(e().mark((function t(n){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return console.log("Testing insert to doc*********"),console.log("response*******",n),t.next=4,Word.run(function(){var t=r(e().mark((function t(r){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return console.log("Inside Testing insert to doc*********"),console.log("Inside response*******",n),r.document.body.insertHtml(n,Word.InsertLocation.end),t.next=6,r.sync();case 6:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 4:case"end":return t.stop()}}),t)}))),c.apply(this,arguments)}Office.onReady(function(){var t=r(e().mark((function t(n){var c;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return document.getElementById("chatWindow").innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">'.concat("Hi, I am your word assistant bot-NoviWord","</div>"),t.next=3,a();case 3:c=t.sent,n.host===Office.HostType.Word&&(document.getElementById("askButton").onclick=r(e().mark((function t(){var n;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if(!(n=document.getElementById("userInput").value)){t.next=6;break}return document.getElementById("headerId").style.display="none",o(n,"","User"),t.next=6,s(c,n);case 6:case"end":return t.stop()}}),t)}))),document.getElementById("userInput").addEventListener("keydown",function(){var t=r(e().mark((function t(n){var r;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if("Enter"!==n.key){t.next=8;break}if(n.preventDefault(),document.getElementById("headerId").style.display="none",!(r=document.getElementById("userInput").value)){t.next=8;break}return o(r,"","User"),t.next=8,s(c,r);case 8:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()),document.getElementById("insertButton").onclick=r(e().mark((function t(){var n;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if(!(n=document.getElementById("chatWindow").lastChild?document.getElementById("chatWindow").lastChild.innerHTML:"")){t.next=4;break}return t.next=4,i(n);case 4:case"end":return t.stop()}}),t)}))));case 5:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());var a=function(){var t=r(e().mark((function t(){var n,r,i;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,t.next=3,fetch("https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview");case 3:return n=t.sent,t.next=6,n.json();case 6:if(r=t.sent,(i=new window.DirectLine.DirectLine({token:r.token}))&&i.activity$){t.next=10;break}throw new Error("DirectLine instance failed to initialize");case 10:return i.postActivity({from:{id:"10",name:"User"},type:"message",text:"Hi"}).subscribe((function(t){return console.log("Message sent with ID:",t)}),(function(t){return console.error("Error sending message:",t)})),i.activity$.subscribe((function(t){console.log("Testing activity: ",t),console.log("Role",t.from.role),"message"!==t.type||"10"===t.from.id||t.recipient||(console.log("Bot Response: ",t.text),o(!1,t,t.from.role))})),t.abrupt("return",i);case 15:t.prev=15,t.t0=t.catch(0),console.error("Error initializing DirectLine:",t.t0);case 18:case"end":return t.stop()}}),t,null,[[0,15]])})));return function(){return t.apply(this,arguments)}}(),s=function(){var t=r(e().mark((function t(n,r){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:n.postActivity({from:{id:"10",name:"User"},type:"message",text:r}).subscribe((function(t){return console.log("Message sent with ID:",t)}),(function(t){return console.error("Error sending message:",t)}));case 1:case"end":return t.stop()}}),t)})));return function(e,n){return t.apply(this,arguments)}}()}(),function(){"use strict";new URL(n(58394),n.b),new URL(n(84400),n.b),new URL(n(60947),n.b)}()}();
//# sourceMappingURL=taskpane.js.map
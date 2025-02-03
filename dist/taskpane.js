/*! For license information please see taskpane.js.LICENSE.txt */
!function(){var t={84400:function(t,e,r){"use strict";t.exports=r.p+"assets/copilot.png"},60947:function(t,e,r){"use strict";t.exports=r.p+"7191ec5570814d614a18.js"},58394:function(t,e,r){"use strict";t.exports=r.p+"faca41d22060e85c0a8c.css"}},e={};function r(n){var o=e[n];if(void 0!==o)return o.exports;var i=e[n]={exports:{}};return t[n](i,i.exports,r),i.exports}r.m=t,r.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),r.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;r.g.importScripts&&(t=r.g.location+"");var e=r.g.document;if(!t&&e&&(e.currentScript&&"SCRIPT"===e.currentScript.tagName.toUpperCase()&&(t=e.currentScript.src),!t)){var n=e.getElementsByTagName("script");if(n.length)for(var o=n.length-1;o>-1&&(!t||!/^http(s?):/.test(t));)t=n[o--].src}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),r.p=t}(),r.b=document.baseURI||self.location.href,function(){function t(e){return t="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},t(e)}function e(){"use strict";e=function(){return n};var r,n={},o=Object.prototype,i=o.hasOwnProperty,a=Object.defineProperty||function(t,e,r){t[e]=r.value},c="function"==typeof Symbol?Symbol:{},s=c.iterator||"@@iterator",u=c.asyncIterator||"@@asyncIterator",l=c.toStringTag||"@@toStringTag";function f(t,e,r){return Object.defineProperty(t,e,{value:r,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{f({},"")}catch(r){f=function(t,e,r){return t[e]=r}}function p(t,e,r,n){var o=e&&e.prototype instanceof w?e:w,i=Object.create(o.prototype),c=new N(n||[]);return a(i,"_invoke",{value:j(t,r,c)}),i}function h(t,e,r){try{return{type:"normal",arg:t.call(e,r)}}catch(t){return{type:"throw",arg:t}}}n.wrap=p;var d="suspendedStart",v="suspendedYield",y="executing",m="completed",g={};function w(){}function b(){}function x(){}var E={};f(E,s,(function(){return this}));var L=Object.getPrototypeOf,k=L&&L(L(P([])));k&&k!==o&&i.call(k,s)&&(E=k);var I=x.prototype=w.prototype=Object.create(E);function T(t){["next","throw","return"].forEach((function(e){f(t,e,(function(t){return this._invoke(e,t)}))}))}function B(e,r){function n(o,a,c,s){var u=h(e[o],e,a);if("throw"!==u.type){var l=u.arg,f=l.value;return f&&"object"==t(f)&&i.call(f,"__await")?r.resolve(f.__await).then((function(t){n("next",t,c,s)}),(function(t){n("throw",t,c,s)})):r.resolve(f).then((function(t){l.value=t,c(l)}),(function(t){return n("throw",t,c,s)}))}s(u.arg)}var o;a(this,"_invoke",{value:function(t,e){function i(){return new r((function(r,o){n(t,e,r,o)}))}return o=o?o.then(i,i):i()}})}function j(t,e,n){var o=d;return function(i,a){if(o===y)throw Error("Generator is already running");if(o===m){if("throw"===i)throw a;return{value:r,done:!0}}for(n.method=i,n.arg=a;;){var c=n.delegate;if(c){var s=O(c,n);if(s){if(s===g)continue;return s}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if(o===d)throw o=m,n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);o=y;var u=h(t,e,n);if("normal"===u.type){if(o=n.done?m:v,u.arg===g)continue;return{value:u.arg,done:n.done}}"throw"===u.type&&(o=m,n.method="throw",n.arg=u.arg)}}}function O(t,e){var n=e.method,o=t.iterator[n];if(o===r)return e.delegate=null,"throw"===n&&t.iterator.return&&(e.method="return",e.arg=r,O(t,e),"throw"===e.method)||"return"!==n&&(e.method="throw",e.arg=new TypeError("The iterator does not provide a '"+n+"' method")),g;var i=h(o,t.iterator,e.arg);if("throw"===i.type)return e.method="throw",e.arg=i.arg,e.delegate=null,g;var a=i.arg;return a?a.done?(e[t.resultName]=a.value,e.next=t.nextLoc,"return"!==e.method&&(e.method="next",e.arg=r),e.delegate=null,g):a:(e.method="throw",e.arg=new TypeError("iterator result is not an object"),e.delegate=null,g)}function S(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function _(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function N(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(S,this),this.reset(!0)}function P(e){if(e||""===e){var n=e[s];if(n)return n.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var o=-1,a=function t(){for(;++o<e.length;)if(i.call(e,o))return t.value=e[o],t.done=!1,t;return t.value=r,t.done=!0,t};return a.next=a}}throw new TypeError(t(e)+" is not iterable")}return b.prototype=x,a(I,"constructor",{value:x,configurable:!0}),a(x,"constructor",{value:b,configurable:!0}),b.displayName=f(x,l,"GeneratorFunction"),n.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===b||"GeneratorFunction"===(e.displayName||e.name))},n.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,x):(t.__proto__=x,f(t,l,"GeneratorFunction")),t.prototype=Object.create(I),t},n.awrap=function(t){return{__await:t}},T(B.prototype),f(B.prototype,u,(function(){return this})),n.AsyncIterator=B,n.async=function(t,e,r,o,i){void 0===i&&(i=Promise);var a=new B(p(t,e,r,o),i);return n.isGeneratorFunction(e)?a:a.next().then((function(t){return t.done?t.value:a.next()}))},T(I),f(I,l,"Generator"),f(I,s,(function(){return this})),f(I,"toString",(function(){return"[object Generator]"})),n.keys=function(t){var e=Object(t),r=[];for(var n in e)r.push(n);return r.reverse(),function t(){for(;r.length;){var n=r.pop();if(n in e)return t.value=n,t.done=!1,t}return t.done=!0,t}},n.values=P,N.prototype={constructor:N,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=r,this.done=!1,this.delegate=null,this.method="next",this.arg=r,this.tryEntries.forEach(_),!t)for(var e in this)"t"===e.charAt(0)&&i.call(this,e)&&!isNaN(+e.slice(1))&&(this[e]=r)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var e=this;function n(n,o){return c.type="throw",c.arg=t,e.next=n,o&&(e.method="next",e.arg=r),!!o}for(var o=this.tryEntries.length-1;o>=0;--o){var a=this.tryEntries[o],c=a.completion;if("root"===a.tryLoc)return n("end");if(a.tryLoc<=this.prev){var s=i.call(a,"catchLoc"),u=i.call(a,"finallyLoc");if(s&&u){if(this.prev<a.catchLoc)return n(a.catchLoc,!0);if(this.prev<a.finallyLoc)return n(a.finallyLoc)}else if(s){if(this.prev<a.catchLoc)return n(a.catchLoc,!0)}else{if(!u)throw Error("try statement without catch or finally");if(this.prev<a.finallyLoc)return n(a.finallyLoc)}}}},abrupt:function(t,e){for(var r=this.tryEntries.length-1;r>=0;--r){var n=this.tryEntries[r];if(n.tryLoc<=this.prev&&i.call(n,"finallyLoc")&&this.prev<n.finallyLoc){var o=n;break}}o&&("break"===t||"continue"===t)&&o.tryLoc<=e&&e<=o.finallyLoc&&(o=null);var a=o?o.completion:{};return a.type=t,a.arg=e,o?(this.method="next",this.next=o.finallyLoc,g):this.complete(a)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),g},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.finallyLoc===t)return this.complete(r.completion,r.afterLoc),_(r),g}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.tryLoc===t){var n=r.completion;if("throw"===n.type){var o=n.arg;_(r)}return o}}throw Error("illegal catch attempt")},delegateYield:function(t,e,n){return this.delegate={iterator:P(t),resultName:e,nextLoc:n},"next"===this.method&&(this.arg=r),g}},n}function r(t,e,r,n,o,i,a){try{var c=t[i](a),s=c.value}catch(t){return void r(t)}c.done?e(s):Promise.resolve(s).then(n,o)}function n(t){return function(){var e=this,n=arguments;return new Promise((function(o,i){var a=t.apply(e,n);function c(t){r(a,o,i,c,s,"next",t)}function s(t){r(a,o,i,c,s,"throw",t)}c(void 0)}))}}function o(t,e,r){var n=document.getElementById("chatWindow");e&&e.attachments&&e.attachments.length>0?e.attachments.forEach((function(t){t.content&&t.content.buttons&&t.content.buttons.length>0&&t.content.buttons.forEach((function(e){if("signin"===e.type){var r=document.createElement("button");r.innerText=e.title||"Sign In",r.classList.add("ms-Button","ms-Button--primary"),r.onclick=function(){window.open(e.value,"_blank")},n.innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">'.concat(t.content.text,"</div>"),n.appendChild(r)}}))})):"bot"===r?e.text&&(n.innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">'.concat(e.text,"</div>")):t&&(n.innerHTML+='<div class="user-wrapper">You</div><div class="message user">'.concat(t,"</div>")),function(){var t=document.getElementById("chatWindow");setTimeout((function(){t.scrollTop=t.scrollHeight}),100)}(),document.getElementById("userInput").value=""}function i(t){return a.apply(this,arguments)}function a(){return a=n(e().mark((function t(r){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=n(e().mark((function t(n){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n.document.body.insertText(r,Word.InsertLocation.end),t.next=4,n.sync();case 4:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:case"end":return t.stop()}}),t)}))),a.apply(this,arguments)}Office.onReady(function(){var t=n(e().mark((function t(r){var a;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return document.getElementById("chatWindow").innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">'.concat("Hi, I am your word assistant bot-NoviWord","</div>"),t.next=3,c();case 3:a=t.sent,r.host===Office.HostType.Word&&(document.getElementById("askButton").onclick=n(e().mark((function t(){var r;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if(!(r=document.getElementById("userInput").value)){t.next=6;break}return document.getElementById("headerId").style.display="none",o(r,"","User"),t.next=6,s(a,r);case 6:case"end":return t.stop()}}),t)}))),document.getElementById("userInput").addEventListener("keydown",function(){var t=n(e().mark((function t(r){var n;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if("Enter"!==r.key){t.next=8;break}if(r.preventDefault(),document.getElementById("headerId").style.display="none",!(n=document.getElementById("userInput").value)){t.next=8;break}return o(n,"","User"),t.next=8,s(a,n);case 8:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()),document.getElementById("insertButton").onclick=n(e().mark((function t(){var r;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if(!(r=document.getElementById("chatWindow").lastChild?document.getElementById("chatWindow").lastChild.innerText:"")){t.next=4;break}return t.next=4,i(r);case 4:case"end":return t.stop()}}),t)}))));case 5:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());var c=function(){var t=n(e().mark((function t(){var r,n,i;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,t.next=3,fetch("https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview");case 3:return r=t.sent,t.next=6,r.json();case 6:if(n=t.sent,(i=new window.DirectLine.DirectLine({token:n.token}))&&i.activity$){t.next=10;break}throw new Error("DirectLine instance failed to initialize");case 10:return i.postActivity({from:{id:"10",name:"User"},type:"message",text:"Hi"}).subscribe((function(t){return console.log("Message sent with ID:",t)}),(function(t){return console.error("Error sending message:",t)})),i.activity$.subscribe((function(t){console.log("Testing activity: ",t),console.log("Role",t.from.role),"message"!==t.type||"10"===t.from.id||t.recipient||(console.log("Bot Response: ",t.text),o(!1,t,t.from.role))})),t.abrupt("return",i);case 15:t.prev=15,t.t0=t.catch(0),console.error("Error initializing DirectLine:",t.t0);case 18:case"end":return t.stop()}}),t,null,[[0,15]])})));return function(){return t.apply(this,arguments)}}(),s=function(){var t=n(e().mark((function t(r,n){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:r.postActivity({from:{id:"10",name:"User"},type:"message",text:n}).subscribe((function(t){return console.log("Message sent with ID:",t)}),(function(t){return console.error("Error sending message:",t)}));case 1:case"end":return t.stop()}}),t)})));return function(e,r){return t.apply(this,arguments)}}()}(),function(){"use strict";new URL(r(58394),r.b),new URL(r(84400),r.b),new URL(r(60947),r.b)}()}();
//# sourceMappingURL=taskpane.js.map
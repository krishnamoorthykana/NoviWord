/*! For license information please see 2a916100ee143215a133.js.LICENSE.txt */
function _typeof(t){return _typeof="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},_typeof(t)}function _regeneratorRuntime(){"use strict";_regeneratorRuntime=function(){return e};var t,e={},n=Object.prototype,r=n.hasOwnProperty,o=Object.defineProperty||function(t,e,n){t[e]=n.value},i="function"==typeof Symbol?Symbol:{},a=i.iterator||"@@iterator",c=i.asyncIterator||"@@asyncIterator",s=i.toStringTag||"@@toStringTag";function u(t,e,n){return Object.defineProperty(t,e,{value:n,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{u({},"")}catch(t){u=function(t,e,n){return t[e]=n}}function l(t,e,n,r){var i=e&&e.prototype instanceof g?e:g,a=Object.create(i.prototype),c=new G(r||[]);return o(a,"_invoke",{value:k(t,n,c)}),a}function f(t,e,n){try{return{type:"normal",arg:t.call(e,n)}}catch(t){return{type:"throw",arg:t}}}e.wrap=l;var h="suspendedStart",p="suspendedYield",y="executing",v="completed",d={};function g(){}function m(){}function w(){}var b={};u(b,a,(function(){return this}));var _=Object.getPrototypeOf,x=_&&_(_(O([])));x&&x!==n&&r.call(x,a)&&(b=x);var L=w.prototype=g.prototype=Object.create(b);function E(t){["next","throw","return"].forEach((function(e){u(t,e,(function(t){return this._invoke(e,t)}))}))}function R(t,e){function n(o,i,a,c){var s=f(t[o],t,i);if("throw"!==s.type){var u=s.arg,l=u.value;return l&&"object"==_typeof(l)&&r.call(l,"__await")?e.resolve(l.__await).then((function(t){n("next",t,a,c)}),(function(t){n("throw",t,a,c)})):e.resolve(l).then((function(t){u.value=t,a(u)}),(function(t){return n("throw",t,a,c)}))}c(s.arg)}var i;o(this,"_invoke",{value:function(t,r){function o(){return new e((function(e,o){n(t,r,e,o)}))}return i=i?i.then(o,o):o()}})}function k(e,n,r){var o=h;return function(i,a){if(o===y)throw Error("Generator is already running");if(o===v){if("throw"===i)throw a;return{value:t,done:!0}}for(r.method=i,r.arg=a;;){var c=r.delegate;if(c){var s=T(c,r);if(s){if(s===d)continue;return s}}if("next"===r.method)r.sent=r._sent=r.arg;else if("throw"===r.method){if(o===h)throw o=v,r.arg;r.dispatchException(r.arg)}else"return"===r.method&&r.abrupt("return",r.arg);o=y;var u=f(e,n,r);if("normal"===u.type){if(o=r.done?v:p,u.arg===d)continue;return{value:u.arg,done:r.done}}"throw"===u.type&&(o=v,r.method="throw",r.arg=u.arg)}}}function T(e,n){var r=n.method,o=e.iterator[r];if(o===t)return n.delegate=null,"throw"===r&&e.iterator.return&&(n.method="return",n.arg=t,T(e,n),"throw"===n.method)||"return"!==r&&(n.method="throw",n.arg=new TypeError("The iterator does not provide a '"+r+"' method")),d;var i=f(o,e.iterator,n.arg);if("throw"===i.type)return n.method="throw",n.arg=i.arg,n.delegate=null,d;var a=i.arg;return a?a.done?(n[e.resultName]=a.value,n.next=e.nextLoc,"return"!==n.method&&(n.method="next",n.arg=t),n.delegate=null,d):a:(n.method="throw",n.arg=new TypeError("iterator result is not an object"),n.delegate=null,d)}function I(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function D(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function G(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(I,this),this.reset(!0)}function O(e){if(e||""===e){var n=e[a];if(n)return n.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var o=-1,i=function n(){for(;++o<e.length;)if(r.call(e,o))return n.value=e[o],n.done=!1,n;return n.value=t,n.done=!0,n};return i.next=i}}throw new TypeError(_typeof(e)+" is not iterable")}return m.prototype=w,o(L,"constructor",{value:w,configurable:!0}),o(w,"constructor",{value:m,configurable:!0}),m.displayName=u(w,s,"GeneratorFunction"),e.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===m||"GeneratorFunction"===(e.displayName||e.name))},e.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,w):(t.__proto__=w,u(t,s,"GeneratorFunction")),t.prototype=Object.create(L),t},e.awrap=function(t){return{__await:t}},E(R.prototype),u(R.prototype,c,(function(){return this})),e.AsyncIterator=R,e.async=function(t,n,r,o,i){void 0===i&&(i=Promise);var a=new R(l(t,n,r,o),i);return e.isGeneratorFunction(n)?a:a.next().then((function(t){return t.done?t.value:a.next()}))},E(L),u(L,s,"Generator"),u(L,a,(function(){return this})),u(L,"toString",(function(){return"[object Generator]"})),e.keys=function(t){var e=Object(t),n=[];for(var r in e)n.push(r);return n.reverse(),function t(){for(;n.length;){var r=n.pop();if(r in e)return t.value=r,t.done=!1,t}return t.done=!0,t}},e.values=O,G.prototype={constructor:G,reset:function(e){if(this.prev=0,this.next=0,this.sent=this._sent=t,this.done=!1,this.delegate=null,this.method="next",this.arg=t,this.tryEntries.forEach(D),!e)for(var n in this)"t"===n.charAt(0)&&r.call(this,n)&&!isNaN(+n.slice(1))&&(this[n]=t)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(e){if(this.done)throw e;var n=this;function o(r,o){return c.type="throw",c.arg=e,n.next=r,o&&(n.method="next",n.arg=t),!!o}for(var i=this.tryEntries.length-1;i>=0;--i){var a=this.tryEntries[i],c=a.completion;if("root"===a.tryLoc)return o("end");if(a.tryLoc<=this.prev){var s=r.call(a,"catchLoc"),u=r.call(a,"finallyLoc");if(s&&u){if(this.prev<a.catchLoc)return o(a.catchLoc,!0);if(this.prev<a.finallyLoc)return o(a.finallyLoc)}else if(s){if(this.prev<a.catchLoc)return o(a.catchLoc,!0)}else{if(!u)throw Error("try statement without catch or finally");if(this.prev<a.finallyLoc)return o(a.finallyLoc)}}}},abrupt:function(t,e){for(var n=this.tryEntries.length-1;n>=0;--n){var o=this.tryEntries[n];if(o.tryLoc<=this.prev&&r.call(o,"finallyLoc")&&this.prev<o.finallyLoc){var i=o;break}}i&&("break"===t||"continue"===t)&&i.tryLoc<=e&&e<=i.finallyLoc&&(i=null);var a=i?i.completion:{};return a.type=t,a.arg=e,i?(this.method="next",this.next=i.finallyLoc,d):this.complete(a)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),d},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.finallyLoc===t)return this.complete(n.completion,n.afterLoc),D(n),d}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.tryLoc===t){var r=n.completion;if("throw"===r.type){var o=r.arg;D(n)}return o}}throw Error("illegal catch attempt")},delegateYield:function(e,n,r){return this.delegate={iterator:O(e),resultName:n,nextLoc:r},"next"===this.method&&(this.arg=t),d}},e}function asyncGeneratorStep(t,e,n,r,o,i,a){try{var c=t[i](a),s=c.value}catch(t){return void n(t)}c.done?e(s):Promise.resolve(s).then(r,o)}function _asyncToGenerator(t){return function(){var e=this,n=arguments;return new Promise((function(r,o){var i=t.apply(e,n);function a(t){asyncGeneratorStep(i,r,o,a,c,"next",t)}function c(t){asyncGeneratorStep(i,r,o,a,c,"throw",t)}a(void 0)}))}}function getChatbotResponse(t){return _getChatbotResponse.apply(this,arguments)}function _getChatbotResponse(){return(_getChatbotResponse=_asyncToGenerator(_regeneratorRuntime().mark((function t(e){return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return console.log("Testing"),initializeDirectLine(),t.abrupt("return","This is a response to: "+e);case 3:case"end":return t.stop()}}),t)})))).apply(this,arguments)}function displayChatMessage(t,e){var n=document.getElementById("chatWindow");n.innerHTML+='<div class="user">You:</div><div>'.concat(t,"</div>"),n.innerHTML+='<div class="bot">Bot:</div><div>'.concat(e,"</div>"),document.getElementById("userInput").value=""}function insertResponseIntoDocument(t){return _insertResponseIntoDocument.apply(this,arguments)}function _insertResponseIntoDocument(){return _insertResponseIntoDocument=_asyncToGenerator(_regeneratorRuntime().mark((function t(e){return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=_asyncToGenerator(_regeneratorRuntime().mark((function t(n){return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n.document.body.insertText(e,Word.InsertLocation.end),t.next=4,n.sync();case 4:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:case"end":return t.stop()}}),t)}))),_insertResponseIntoDocument.apply(this,arguments)}Office.onReady((function(t){t.host===Office.HostType.Word&&(document.getElementById("askButton").onclick=_asyncToGenerator(_regeneratorRuntime().mark((function t(){var e,n;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if(!(e=document.getElementById("userInput").value)){t.next=6;break}return t.next=4,getChatbotResponse(e);case 4:n=t.sent,displayChatMessage(e,n);case 6:case"end":return t.stop()}}),t)}))),document.getElementById("insertButton").onclick=_asyncToGenerator(_regeneratorRuntime().mark((function t(){var e;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if(!(e=document.getElementById("chatWindow").lastChild?document.getElementById("chatWindow").lastChild.innerText:"")){t.next=4;break}return t.next=4,insertResponseIntoDocument(e);case 4:case"end":return t.stop()}}),t)}))))}));var initializeDirectLine=function(){var t=_asyncToGenerator(_regeneratorRuntime().mark((function t(){var e,n,r;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,t.next=3,fetch("https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview");case 3:return e=t.sent,t.next=6,e.json();case 6:if(n=t.sent,console.log("Testing data token:"+JSON.stringify(n,null,2)),console.log("DirectLine Object:",window.DirectLine),r=new window.DirectLine.DirectLine({token:n.token}),console.log("directLine*******",r),console.log("DirectLine instance:",r),console.log("DirectLine activity$:",r.activity$),r&&r.activity$){t.next=15;break}throw new Error("DirectLine instance failed to initialize");case 15:r.activity$.subscribe((function(t){console.log("Received activity:",t)})),r.connectionStatus$.subscribe((function(t){console.log("DirectLine connection status:",t)})),r.postActivity({from:{id:"user1",name:"User"},type:"message",text:"Hello, bot!"}).subscribe((function(t){return console.log("Message sent with ID:",t)}),(function(t){return console.error("Error sending message:",t)})),r.activity$.subscribe((function(t){console.log("Testing activity: ",t),"message"!==t.type||"10"===t.from.id||t.recipient||console.log("Testing response: ",t.text)})),t.next=24;break;case 21:t.prev=21,t.t0=t.catch(0),console.error("Error initializing DirectLine:",t.t0);case 24:case"end":return t.stop()}}),t,null,[[0,21]])})));return function(){return t.apply(this,arguments)}}();
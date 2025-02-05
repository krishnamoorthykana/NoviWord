/*! For license information please see 61a493835e9fd23e42fe.js.LICENSE.txt */
function _typeof(t){return _typeof="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},_typeof(t)}function _regeneratorRuntime(){"use strict";_regeneratorRuntime=function(){return e};var t,e={},n=Object.prototype,r=n.hasOwnProperty,o=Object.defineProperty||function(t,e,n){t[e]=n.value},i="function"==typeof Symbol?Symbol:{},a=i.iterator||"@@iterator",c=i.asyncIterator||"@@asyncIterator",s=i.toStringTag||"@@toStringTag";function u(t,e,n){return Object.defineProperty(t,e,{value:n,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{u({},"")}catch(t){u=function(t,e,n){return t[e]=n}}function l(t,e,n,r){var i=e&&e.prototype instanceof m?e:m,a=Object.create(i.prototype),c=new S(r||[]);return o(a,"_invoke",{value:k(t,n,c)}),a}function p(t,e,n){try{return{type:"normal",arg:t.call(e,n)}}catch(t){return{type:"throw",arg:t}}}e.wrap=l;var f="suspendedStart",h="suspendedYield",d="executing",y="completed",g={};function m(){}function v(){}function w(){}var b={};u(b,a,(function(){return this}));var x=Object.getPrototypeOf,_=x&&x(x(G([])));_&&_!==n&&r.call(_,a)&&(b=_);var E=w.prototype=m.prototype=Object.create(b);function L(t){["next","throw","return"].forEach((function(e){u(t,e,(function(t){return this._invoke(e,t)}))}))}function I(t,e){function n(o,i,a,c){var s=p(t[o],t,i);if("throw"!==s.type){var u=s.arg,l=u.value;return l&&"object"==_typeof(l)&&r.call(l,"__await")?e.resolve(l.__await).then((function(t){n("next",t,a,c)}),(function(t){n("throw",t,a,c)})):e.resolve(l).then((function(t){u.value=t,a(u)}),(function(t){return n("throw",t,a,c)}))}c(s.arg)}var i;o(this,"_invoke",{value:function(t,r){function o(){return new e((function(e,o){n(t,r,e,o)}))}return i=i?i.then(o,o):o()}})}function k(e,n,r){var o=f;return function(i,a){if(o===d)throw Error("Generator is already running");if(o===y){if("throw"===i)throw a;return{value:t,done:!0}}for(r.method=i,r.arg=a;;){var c=r.delegate;if(c){var s=T(c,r);if(s){if(s===g)continue;return s}}if("next"===r.method)r.sent=r._sent=r.arg;else if("throw"===r.method){if(o===f)throw o=y,r.arg;r.dispatchException(r.arg)}else"return"===r.method&&r.abrupt("return",r.arg);o=d;var u=p(e,n,r);if("normal"===u.type){if(o=r.done?y:h,u.arg===g)continue;return{value:u.arg,done:r.done}}"throw"===u.type&&(o=y,r.method="throw",r.arg=u.arg)}}}function T(e,n){var r=n.method,o=e.iterator[r];if(o===t)return n.delegate=null,"throw"===r&&e.iterator.return&&(n.method="return",n.arg=t,T(e,n),"throw"===n.method)||"return"!==r&&(n.method="throw",n.arg=new TypeError("The iterator does not provide a '"+r+"' method")),g;var i=p(o,e.iterator,n.arg);if("throw"===i.type)return n.method="throw",n.arg=i.arg,n.delegate=null,g;var a=i.arg;return a?a.done?(n[e.resultName]=a.value,n.next=e.nextLoc,"return"!==n.method&&(n.method="next",n.arg=t),n.delegate=null,g):a:(n.method="throw",n.arg=new TypeError("iterator result is not an object"),n.delegate=null,g)}function R(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function B(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function S(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(R,this),this.reset(!0)}function G(e){if(e||""===e){var n=e[a];if(n)return n.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var o=-1,i=function n(){for(;++o<e.length;)if(r.call(e,o))return n.value=e[o],n.done=!1,n;return n.value=t,n.done=!0,n};return i.next=i}}throw new TypeError(_typeof(e)+" is not iterable")}return v.prototype=w,o(E,"constructor",{value:w,configurable:!0}),o(w,"constructor",{value:v,configurable:!0}),v.displayName=u(w,s,"GeneratorFunction"),e.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===v||"GeneratorFunction"===(e.displayName||e.name))},e.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,w):(t.__proto__=w,u(t,s,"GeneratorFunction")),t.prototype=Object.create(E),t},e.awrap=function(t){return{__await:t}},L(I.prototype),u(I.prototype,c,(function(){return this})),e.AsyncIterator=I,e.async=function(t,n,r,o,i){void 0===i&&(i=Promise);var a=new I(l(t,n,r,o),i);return e.isGeneratorFunction(n)?a:a.next().then((function(t){return t.done?t.value:a.next()}))},L(E),u(E,s,"Generator"),u(E,a,(function(){return this})),u(E,"toString",(function(){return"[object Generator]"})),e.keys=function(t){var e=Object(t),n=[];for(var r in e)n.push(r);return n.reverse(),function t(){for(;n.length;){var r=n.pop();if(r in e)return t.value=r,t.done=!1,t}return t.done=!0,t}},e.values=G,S.prototype={constructor:S,reset:function(e){if(this.prev=0,this.next=0,this.sent=this._sent=t,this.done=!1,this.delegate=null,this.method="next",this.arg=t,this.tryEntries.forEach(B),!e)for(var n in this)"t"===n.charAt(0)&&r.call(this,n)&&!isNaN(+n.slice(1))&&(this[n]=t)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(e){if(this.done)throw e;var n=this;function o(r,o){return c.type="throw",c.arg=e,n.next=r,o&&(n.method="next",n.arg=t),!!o}for(var i=this.tryEntries.length-1;i>=0;--i){var a=this.tryEntries[i],c=a.completion;if("root"===a.tryLoc)return o("end");if(a.tryLoc<=this.prev){var s=r.call(a,"catchLoc"),u=r.call(a,"finallyLoc");if(s&&u){if(this.prev<a.catchLoc)return o(a.catchLoc,!0);if(this.prev<a.finallyLoc)return o(a.finallyLoc)}else if(s){if(this.prev<a.catchLoc)return o(a.catchLoc,!0)}else{if(!u)throw Error("try statement without catch or finally");if(this.prev<a.finallyLoc)return o(a.finallyLoc)}}}},abrupt:function(t,e){for(var n=this.tryEntries.length-1;n>=0;--n){var o=this.tryEntries[n];if(o.tryLoc<=this.prev&&r.call(o,"finallyLoc")&&this.prev<o.finallyLoc){var i=o;break}}i&&("break"===t||"continue"===t)&&i.tryLoc<=e&&e<=i.finallyLoc&&(i=null);var a=i?i.completion:{};return a.type=t,a.arg=e,i?(this.method="next",this.next=i.finallyLoc,g):this.complete(a)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),g},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.finallyLoc===t)return this.complete(n.completion,n.afterLoc),B(n),g}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.tryLoc===t){var r=n.completion;if("throw"===r.type){var o=r.arg;B(n)}return o}}throw Error("illegal catch attempt")},delegateYield:function(e,n,r){return this.delegate={iterator:G(e),resultName:n,nextLoc:r},"next"===this.method&&(this.arg=t),g}},e}function asyncGeneratorStep(t,e,n,r,o,i,a){try{var c=t[i](a),s=c.value}catch(t){return void n(t)}c.done?e(s):Promise.resolve(s).then(r,o)}function _asyncToGenerator(t){return function(){var e=this,n=arguments;return new Promise((function(r,o){var i=t.apply(e,n);function a(t){asyncGeneratorStep(i,r,o,a,c,"next",t)}function c(t){asyncGeneratorStep(i,r,o,a,c,"throw",t)}a(void 0)}))}}var speechFlag=!1;function displayStartingMessage(t){document.getElementById("chatWindow").innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">'.concat(t,"</div>")}function displayChatMessage(t,e,n){var r=document.getElementById("chatWindow");e&&e.attachments&&e.attachments.length>0?e.attachments.forEach((function(t){t.content&&t.content.buttons&&t.content.buttons.length>0&&t.content.buttons.forEach((function(e){if("signin"===e.type){var n=document.createElement("button");n.innerText=e.title||"Sign In",n.classList.add("ms-Button","ms-Button--primary"),n.onclick=function(){window.open(e.value,"_blank")},r.innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">'.concat(t.content.text,"</div>"),r.appendChild(n)}}))})):"bot"===n?"Generate"===e.speak?(insertResponseIntoDocument(e.text),r.innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">SOW content generated in document</div>'):e.text&&(r.innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">'.concat(e.text,"</div>"),speechFlag&&(speakText(e.text),speechFlag=!1)):t&&(r.innerHTML+='<div class="user-wrapper">You</div><div class="message user">'.concat(t,"</div>")),scrollToBottom(),document.getElementById("userInput").value=""}function insertResponseIntoDocument(t){return _insertResponseIntoDocument.apply(this,arguments)}function _insertResponseIntoDocument(){return _insertResponseIntoDocument=_asyncToGenerator(_regeneratorRuntime().mark((function t(e){return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return console.log("Testing insert to doc*********"),console.log("response*******",e),t.next=4,Word.run(function(){var t=_asyncToGenerator(_regeneratorRuntime().mark((function t(n){return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return console.log("Inside Testing insert to doc*********"),console.log("Inside response*******",e),n.document.body.insertHtml(e,Word.InsertLocation.end),t.next=6,n.sync();case 6:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 4:case"end":return t.stop()}}),t)}))),_insertResponseIntoDocument.apply(this,arguments)}Office.onReady(function(){var t=_asyncToGenerator(_regeneratorRuntime().mark((function t(e){var n;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return displayStartingMessage("Hi, I am your word assistant bot-NoviWord"),t.next=3,initializeDirectLine();case 3:n=t.sent,e.host===Office.HostType.Word&&(document.getElementById("askButton").onclick=_asyncToGenerator(_regeneratorRuntime().mark((function t(){var e;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if(!(e=document.getElementById("userInput").value)){t.next=6;break}return document.getElementById("headerId").style.display="none",displayChatMessage(e,"","User"),t.next=6,getBotResponse(n,e);case 6:case"end":return t.stop()}}),t)}))),document.getElementById("userInput").addEventListener("keydown",function(){var t=_asyncToGenerator(_regeneratorRuntime().mark((function t(e){var r;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if("Enter"!==e.key){t.next=8;break}if(e.preventDefault(),document.getElementById("headerId").style.display="none",!(r=document.getElementById("userInput").value)){t.next=8;break}return displayChatMessage(r,"","User"),t.next=8,getBotResponse(n,r);case 8:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()),document.getElementById("insertButton").onclick=_asyncToGenerator(_regeneratorRuntime().mark((function t(){var e;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if(!(e=document.getElementById("chatWindow").lastChild?document.getElementById("chatWindow").lastChild.innerHTML:"")){t.next=4;break}return t.next=4,insertResponseIntoDocument(e);case 4:case"end":return t.stop()}}),t)}))));case 5:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());var initializeDirectLine=function(){var t=_asyncToGenerator(_regeneratorRuntime().mark((function t(){var e,n,r;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,t.next=3,fetch("https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview");case 3:return e=t.sent,t.next=6,e.json();case 6:if(n=t.sent,(r=new window.DirectLine.DirectLine({token:n.token}))&&r.activity$){t.next=10;break}throw new Error("DirectLine instance failed to initialize");case 10:return r.postActivity({from:{id:"10",name:"User"},type:"message",text:"Hi"}).subscribe((function(t){return console.log("Message sent with ID:",t)}),(function(t){return console.error("Error sending message:",t)})),r.activity$.subscribe((function(t){console.log("Testing activity: ",t),console.log("Role",t.from.role),"message"!==t.type||"10"===t.from.id||t.recipient||(console.log("Bot Response: ",t.text),displayChatMessage(!1,t,t.from.role))})),t.abrupt("return",r);case 15:t.prev=15,t.t0=t.catch(0),console.error("Error initializing DirectLine:",t.t0);case 18:case"end":return t.stop()}}),t,null,[[0,15]])})));return function(){return t.apply(this,arguments)}}(),getBotResponse=function(){var t=_asyncToGenerator(_regeneratorRuntime().mark((function t(e,n){return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:e.postActivity({from:{id:"10",name:"User"},type:"message",text:n}).subscribe((function(t){return console.log("Message sent with ID:",t)}),(function(t){return console.error("Error sending message:",t)}));case 1:case"end":return t.stop()}}),t)})));return function(e,n){return t.apply(this,arguments)}}();function scrollToBottom(){var t=document.getElementById("chatWindow");setTimeout((function(){t.scrollTop=t.scrollHeight}),100)}function speakText(t){console.log("Testing Text to Speech");var e=new SpeechSynthesisUtterance(t);e.lang="en-US",e.rate=1,e.pitch=1,e.volume=1,window.speechSynthesis.speak(e)}document.getElementById("startSpeechButton").addEventListener("click",(function(){window.open("speech.html","SpeechRecognition","width=400,height=300"),speechFlag=!0,window.addEventListener("message",(function(t){if(t.origin===window.location.origin){var e=t.data;console.log(e),document.getElementById("userInput").value=e}}))}));
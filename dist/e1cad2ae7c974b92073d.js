/*! For license information please see e1cad2ae7c974b92073d.js.LICENSE.txt */
function _typeof(e){return _typeof="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(e){return typeof e}:function(e){return e&&"function"==typeof Symbol&&e.constructor===Symbol&&e!==Symbol.prototype?"symbol":typeof e},_typeof(e)}function _regeneratorRuntime(){"use strict";_regeneratorRuntime=function(){return t};var e,t={},n=Object.prototype,r=n.hasOwnProperty,o=Object.defineProperty||function(e,t,n){e[t]=n.value},i="function"==typeof Symbol?Symbol:{},a=i.iterator||"@@iterator",c=i.asyncIterator||"@@asyncIterator",s=i.toStringTag||"@@toStringTag";function u(e,t,n){return Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}),e[t]}try{u({},"")}catch(e){u=function(e,t,n){return e[t]=n}}function l(e,t,n,r){var i=t&&t.prototype instanceof m?t:m,a=Object.create(i.prototype),c=new S(r||[]);return o(a,"_invoke",{value:R(e,n,c)}),a}function f(e,t,n){try{return{type:"normal",arg:e.call(t,n)}}catch(e){return{type:"throw",arg:e}}}t.wrap=l;var p="suspendedStart",h="suspendedYield",d="executing",g="completed",y={};function m(){}function v(){}function w(){}var b={};u(b,a,(function(){return this}));var x=Object.getPrototypeOf,_=x&&x(x(G([])));_&&_!==n&&r.call(_,a)&&(b=_);var E=w.prototype=m.prototype=Object.create(b);function L(e){["next","throw","return"].forEach((function(t){u(e,t,(function(e){return this._invoke(t,e)}))}))}function I(e,t){function n(o,i,a,c){var s=f(e[o],e,i);if("throw"!==s.type){var u=s.arg,l=u.value;return l&&"object"==_typeof(l)&&r.call(l,"__await")?t.resolve(l.__await).then((function(e){n("next",e,a,c)}),(function(e){n("throw",e,a,c)})):t.resolve(l).then((function(e){u.value=e,a(u)}),(function(e){return n("throw",e,a,c)}))}c(s.arg)}var i;o(this,"_invoke",{value:function(e,r){function o(){return new t((function(t,o){n(e,r,t,o)}))}return i=i?i.then(o,o):o()}})}function R(t,n,r){var o=p;return function(i,a){if(o===d)throw Error("Generator is already running");if(o===g){if("throw"===i)throw a;return{value:e,done:!0}}for(r.method=i,r.arg=a;;){var c=r.delegate;if(c){var s=k(c,r);if(s){if(s===y)continue;return s}}if("next"===r.method)r.sent=r._sent=r.arg;else if("throw"===r.method){if(o===p)throw o=g,r.arg;r.dispatchException(r.arg)}else"return"===r.method&&r.abrupt("return",r.arg);o=d;var u=f(t,n,r);if("normal"===u.type){if(o=r.done?g:h,u.arg===y)continue;return{value:u.arg,done:r.done}}"throw"===u.type&&(o=g,r.method="throw",r.arg=u.arg)}}}function k(t,n){var r=n.method,o=t.iterator[r];if(o===e)return n.delegate=null,"throw"===r&&t.iterator.return&&(n.method="return",n.arg=e,k(t,n),"throw"===n.method)||"return"!==r&&(n.method="throw",n.arg=new TypeError("The iterator does not provide a '"+r+"' method")),y;var i=f(o,t.iterator,n.arg);if("throw"===i.type)return n.method="throw",n.arg=i.arg,n.delegate=null,y;var a=i.arg;return a?a.done?(n[t.resultName]=a.value,n.next=t.nextLoc,"return"!==n.method&&(n.method="next",n.arg=e),n.delegate=null,y):a:(n.method="throw",n.arg=new TypeError("iterator result is not an object"),n.delegate=null,y)}function T(e){var t={tryLoc:e[0]};1 in e&&(t.catchLoc=e[1]),2 in e&&(t.finallyLoc=e[2],t.afterLoc=e[3]),this.tryEntries.push(t)}function B(e){var t=e.completion||{};t.type="normal",delete t.arg,e.completion=t}function S(e){this.tryEntries=[{tryLoc:"root"}],e.forEach(T,this),this.reset(!0)}function G(t){if(t||""===t){var n=t[a];if(n)return n.call(t);if("function"==typeof t.next)return t;if(!isNaN(t.length)){var o=-1,i=function n(){for(;++o<t.length;)if(r.call(t,o))return n.value=t[o],n.done=!1,n;return n.value=e,n.done=!0,n};return i.next=i}}throw new TypeError(_typeof(t)+" is not iterable")}return v.prototype=w,o(E,"constructor",{value:w,configurable:!0}),o(w,"constructor",{value:v,configurable:!0}),v.displayName=u(w,s,"GeneratorFunction"),t.isGeneratorFunction=function(e){var t="function"==typeof e&&e.constructor;return!!t&&(t===v||"GeneratorFunction"===(t.displayName||t.name))},t.mark=function(e){return Object.setPrototypeOf?Object.setPrototypeOf(e,w):(e.__proto__=w,u(e,s,"GeneratorFunction")),e.prototype=Object.create(E),e},t.awrap=function(e){return{__await:e}},L(I.prototype),u(I.prototype,c,(function(){return this})),t.AsyncIterator=I,t.async=function(e,n,r,o,i){void 0===i&&(i=Promise);var a=new I(l(e,n,r,o),i);return t.isGeneratorFunction(n)?a:a.next().then((function(e){return e.done?e.value:a.next()}))},L(E),u(E,s,"Generator"),u(E,a,(function(){return this})),u(E,"toString",(function(){return"[object Generator]"})),t.keys=function(e){var t=Object(e),n=[];for(var r in t)n.push(r);return n.reverse(),function e(){for(;n.length;){var r=n.pop();if(r in t)return e.value=r,e.done=!1,e}return e.done=!0,e}},t.values=G,S.prototype={constructor:S,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=e,this.done=!1,this.delegate=null,this.method="next",this.arg=e,this.tryEntries.forEach(B),!t)for(var n in this)"t"===n.charAt(0)&&r.call(this,n)&&!isNaN(+n.slice(1))&&(this[n]=e)},stop:function(){this.done=!0;var e=this.tryEntries[0].completion;if("throw"===e.type)throw e.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var n=this;function o(r,o){return c.type="throw",c.arg=t,n.next=r,o&&(n.method="next",n.arg=e),!!o}for(var i=this.tryEntries.length-1;i>=0;--i){var a=this.tryEntries[i],c=a.completion;if("root"===a.tryLoc)return o("end");if(a.tryLoc<=this.prev){var s=r.call(a,"catchLoc"),u=r.call(a,"finallyLoc");if(s&&u){if(this.prev<a.catchLoc)return o(a.catchLoc,!0);if(this.prev<a.finallyLoc)return o(a.finallyLoc)}else if(s){if(this.prev<a.catchLoc)return o(a.catchLoc,!0)}else{if(!u)throw Error("try statement without catch or finally");if(this.prev<a.finallyLoc)return o(a.finallyLoc)}}}},abrupt:function(e,t){for(var n=this.tryEntries.length-1;n>=0;--n){var o=this.tryEntries[n];if(o.tryLoc<=this.prev&&r.call(o,"finallyLoc")&&this.prev<o.finallyLoc){var i=o;break}}i&&("break"===e||"continue"===e)&&i.tryLoc<=t&&t<=i.finallyLoc&&(i=null);var a=i?i.completion:{};return a.type=e,a.arg=t,i?(this.method="next",this.next=i.finallyLoc,y):this.complete(a)},complete:function(e,t){if("throw"===e.type)throw e.arg;return"break"===e.type||"continue"===e.type?this.next=e.arg:"return"===e.type?(this.rval=this.arg=e.arg,this.method="return",this.next="end"):"normal"===e.type&&t&&(this.next=t),y},finish:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var n=this.tryEntries[t];if(n.finallyLoc===e)return this.complete(n.completion,n.afterLoc),B(n),y}},catch:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var n=this.tryEntries[t];if(n.tryLoc===e){var r=n.completion;if("throw"===r.type){var o=r.arg;B(n)}return o}}throw Error("illegal catch attempt")},delegateYield:function(t,n,r){return this.delegate={iterator:G(t),resultName:n,nextLoc:r},"next"===this.method&&(this.arg=e),y}},t}function asyncGeneratorStep(e,t,n,r,o,i,a){try{var c=e[i](a),s=c.value}catch(e){return void n(e)}c.done?t(s):Promise.resolve(s).then(r,o)}function _asyncToGenerator(e){return function(){var t=this,n=arguments;return new Promise((function(r,o){var i=e.apply(t,n);function a(e){asyncGeneratorStep(i,r,o,a,c,"next",e)}function c(e){asyncGeneratorStep(i,r,o,a,c,"throw",e)}a(void 0)}))}}function displayStartingMessage(e){document.getElementById("chatWindow").innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">'.concat(e,"</div>")}function displayChatMessage(e,t,n){var r=document.getElementById("chatWindow");t&&t.attachments&&t.attachments.length>0?t.attachments.forEach((function(e){e.content&&e.content.buttons&&e.content.buttons.length>0&&e.content.buttons.forEach((function(t){if("signin"===t.type){var n=document.createElement("button");n.innerText=t.title||"Sign In",n.classList.add("ms-Button","ms-Button--primary"),n.onclick=function(){window.open(t.value,"_blank")},r.innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">'.concat(e.content.text,"</div>"),r.appendChild(n)}}))})):"bot"===n?"Generate"===t.speak?(insertResponseIntoDocument(t.text),r.innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviWord</div><div class="message bot">SOW content generated in document</div>'):t.text&&(r.innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviWord</div><div class="message bot">'.concat(t.text,"</div>")):e&&(r.innerHTML+='<div class="user-wrapper">You</div><div class="message user">'.concat(e,"</div>")),scrollToBottom(),document.getElementById("userInput").value=""}function insertResponseIntoDocument(e){return _insertResponseIntoDocument.apply(this,arguments)}function _insertResponseIntoDocument(){return _insertResponseIntoDocument=_asyncToGenerator(_regeneratorRuntime().mark((function e(t){return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return console.log("Testing insert to doc*********"),console.log("response*******",t),e.next=4,Word.run(function(){var e=_asyncToGenerator(_regeneratorRuntime().mark((function e(n){return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return console.log("Inside Testing insert to doc*********"),console.log("Inside response*******",t),n.document.body.insertHtml(t,Word.InsertLocation.end),e.next=6,n.sync();case 6:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 4:case"end":return e.stop()}}),e)}))),_insertResponseIntoDocument.apply(this,arguments)}Office.onReady(function(){var e=_asyncToGenerator(_regeneratorRuntime().mark((function e(t){var n;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return displayStartingMessage("Hi, I am your word assistant bot-NoviWord"),e.next=3,initializeDirectLine();case 3:n=e.sent,t.host===Office.HostType.Word&&(document.getElementById("askButton").onclick=_asyncToGenerator(_regeneratorRuntime().mark((function e(){var t;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if(!(t=document.getElementById("userInput").value)){e.next=6;break}return document.getElementById("headerId").style.display="none",displayChatMessage(t,"","User"),e.next=6,getBotResponse(n,t);case 6:case"end":return e.stop()}}),e)}))),document.getElementById("userInput").addEventListener("keydown",function(){var e=_asyncToGenerator(_regeneratorRuntime().mark((function e(t){var r;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if("Enter"!==t.key){e.next=8;break}if(t.preventDefault(),document.getElementById("headerId").style.display="none",!(r=document.getElementById("userInput").value)){e.next=8;break}return displayChatMessage(r,"","User"),e.next=8,getBotResponse(n,r);case 8:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}()),document.getElementById("insertButton").onclick=_asyncToGenerator(_regeneratorRuntime().mark((function e(){var t;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if(!(t=document.getElementById("chatWindow").lastChild?document.getElementById("chatWindow").lastChild.innerHTML:"")){e.next=4;break}return e.next=4,insertResponseIntoDocument(t);case 4:case"end":return e.stop()}}),e)}))));case 5:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());var initializeDirectLine=function(){var e=_asyncToGenerator(_regeneratorRuntime().mark((function e(){var t,n,r;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,fetch("https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview");case 3:return t=e.sent,e.next=6,t.json();case 6:if(n=e.sent,(r=new window.DirectLine.DirectLine({token:n.token}))&&r.activity$){e.next=10;break}throw new Error("DirectLine instance failed to initialize");case 10:return r.postActivity({from:{id:"10",name:"User"},type:"message",text:"Hi"}).subscribe((function(e){return console.log("Message sent with ID:",e)}),(function(e){return console.error("Error sending message:",e)})),r.activity$.subscribe((function(e){console.log("Testing activity: ",e),console.log("Role",e.from.role),"message"!==e.type||"10"===e.from.id||e.recipient||(console.log("Bot Response: ",e.text),displayChatMessage(!1,e,e.from.role))})),e.abrupt("return",r);case 15:e.prev=15,e.t0=e.catch(0),console.error("Error initializing DirectLine:",e.t0);case 18:case"end":return e.stop()}}),e,null,[[0,15]])})));return function(){return e.apply(this,arguments)}}(),getBotResponse=function(){var e=_asyncToGenerator(_regeneratorRuntime().mark((function e(t,n){return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:t.postActivity({from:{id:"10",name:"User"},type:"message",text:n}).subscribe((function(e){return console.log("Message sent with ID:",e)}),(function(e){return console.error("Error sending message:",e)}));case 1:case"end":return e.stop()}}),e)})));return function(t,n){return e.apply(this,arguments)}}();function scrollToBottom(){var e=document.getElementById("chatWindow");setTimeout((function(){e.scrollTop=e.scrollHeight}),100)}var recognition=null;function startVoiceInput(){recognition&&recognition.abort&&recognition.abort(),(recognition=new(window.SpeechRecognition||window.webkitSpeechRecognition)).lang="en-US",recognition.interimResults=!1,recognition.maxAlternatives=1,recognition.continuous=!0,recognition.onstart=function(){console.log("Speech recognition started.")},recognition.onresult=function(e){var t=e.results[0][0].transcript;console.log("Recognized text:",t),document.getElementById("userInput").value=t},recognition.onerror=function(e){console.error("Speech recognition error:",e.error),"aborted"===e.error&&(console.log("Restarting recognition..."),recognition.start())},recognition.onend=function(){console.log("Speech recognition ended.")},recognition.start()}document.getElementById("speakButton").addEventListener("click",_asyncToGenerator(_regeneratorRuntime().mark((function e(){return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,navigator.mediaDevices.getUserMedia({audio:!0});case 3:console.log("Microphone access granted."),startVoiceInput(),e.next=10;break;case 7:e.prev=7,e.t0=e.catch(0),console.error("Microphone access denied:",e.t0);case 10:case"end":return e.stop()}}),e,null,[[0,7]])}))));
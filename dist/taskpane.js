/*! For license information please see taskpane.js.LICENSE.txt */
!function(){var e={84400:function(e,t,n){"use strict";e.exports=n.p+"assets/copilot.png"},60947:function(e,t,n){"use strict";e.exports=n.p+"f15beda057dfb8fc652e.js"},58394:function(e,t,n){"use strict";e.exports=n.p+"b3519a760f8b0419f70b.css"}},t={};function n(r){var o=t[r];if(void 0!==o)return o.exports;var c=t[r]={exports:{}};return e[r](c,c.exports,n),c.exports}n.m=e,n.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;n.g.importScripts&&(e=n.g.location+"");var t=n.g.document;if(!e&&t&&(t.currentScript&&"SCRIPT"===t.currentScript.tagName.toUpperCase()&&(e=t.currentScript.src),!e)){var r=t.getElementsByTagName("script");if(r.length)for(var o=r.length-1;o>-1&&(!e||!/^http(s?):/.test(e));)e=r[o--].src}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),n.p=e}(),n.b=document.baseURI||self.location.href,function(){function e(t){return e="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(e){return typeof e}:function(e){return e&&"function"==typeof Symbol&&e.constructor===Symbol&&e!==Symbol.prototype?"symbol":typeof e},e(t)}function t(){"use strict";t=function(){return r};var n,r={},o=Object.prototype,c=o.hasOwnProperty,i=Object.defineProperty||function(e,t,n){e[t]=n.value},a="function"==typeof Symbol?Symbol:{},s=a.iterator||"@@iterator",u=a.asyncIterator||"@@asyncIterator",l=a.toStringTag||"@@toStringTag";function p(e,t,n){return Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}),e[t]}try{p({},"")}catch(n){p=function(e,t,n){return e[t]=n}}function f(e,t,n,r){var o=t&&t.prototype instanceof w?t:w,c=Object.create(o.prototype),a=new j(r||[]);return i(c,"_invoke",{value:B(e,n,a)}),c}function d(e,t,n){try{return{type:"normal",arg:e.call(t,n)}}catch(e){return{type:"throw",arg:e}}}r.wrap=f;var h="suspendedStart",m="suspendedYield",g="executing",v="completed",y={};function w(){}function x(){}function b(){}var k={};p(k,s,(function(){return this}));var L=Object.getPrototypeOf,E=L&&L(L(O([])));E&&E!==o&&c.call(E,s)&&(k=E);var T=b.prototype=w.prototype=Object.create(k);function I(e){["next","throw","return"].forEach((function(t){p(e,t,(function(e){return this._invoke(t,e)}))}))}function S(t,n){function r(o,i,a,s){var u=d(t[o],t,i);if("throw"!==u.type){var l=u.arg,p=l.value;return p&&"object"==e(p)&&c.call(p,"__await")?n.resolve(p.__await).then((function(e){r("next",e,a,s)}),(function(e){r("throw",e,a,s)})):n.resolve(p).then((function(e){l.value=e,a(l)}),(function(e){return r("throw",e,a,s)}))}s(u.arg)}var o;i(this,"_invoke",{value:function(e,t){function c(){return new n((function(n,o){r(e,t,n,o)}))}return o=o?o.then(c,c):c()}})}function B(e,t,r){var o=h;return function(c,i){if(o===g)throw Error("Generator is already running");if(o===v){if("throw"===c)throw i;return{value:n,done:!0}}for(r.method=c,r.arg=i;;){var a=r.delegate;if(a){var s=W(a,r);if(s){if(s===y)continue;return s}}if("next"===r.method)r.sent=r._sent=r.arg;else if("throw"===r.method){if(o===h)throw o=v,r.arg;r.dispatchException(r.arg)}else"return"===r.method&&r.abrupt("return",r.arg);o=g;var u=d(e,t,r);if("normal"===u.type){if(o=r.done?v:m,u.arg===y)continue;return{value:u.arg,done:r.done}}"throw"===u.type&&(o=v,r.method="throw",r.arg=u.arg)}}}function W(e,t){var r=t.method,o=e.iterator[r];if(o===n)return t.delegate=null,"throw"===r&&e.iterator.return&&(t.method="return",t.arg=n,W(e,t),"throw"===t.method)||"return"!==r&&(t.method="throw",t.arg=new TypeError("The iterator does not provide a '"+r+"' method")),y;var c=d(o,e.iterator,t.arg);if("throw"===c.type)return t.method="throw",t.arg=c.arg,t.delegate=null,y;var i=c.arg;return i?i.done?(t[e.resultName]=i.value,t.next=e.nextLoc,"return"!==t.method&&(t.method="next",t.arg=n),t.delegate=null,y):i:(t.method="throw",t.arg=new TypeError("iterator result is not an object"),t.delegate=null,y)}function H(e){var t={tryLoc:e[0]};1 in e&&(t.catchLoc=e[1]),2 in e&&(t.finallyLoc=e[2],t.afterLoc=e[3]),this.tryEntries.push(t)}function N(e){var t=e.completion||{};t.type="normal",delete t.arg,e.completion=t}function j(e){this.tryEntries=[{tryLoc:"root"}],e.forEach(H,this),this.reset(!0)}function O(t){if(t||""===t){var r=t[s];if(r)return r.call(t);if("function"==typeof t.next)return t;if(!isNaN(t.length)){var o=-1,i=function e(){for(;++o<t.length;)if(c.call(t,o))return e.value=t[o],e.done=!1,e;return e.value=n,e.done=!0,e};return i.next=i}}throw new TypeError(e(t)+" is not iterable")}return x.prototype=b,i(T,"constructor",{value:b,configurable:!0}),i(b,"constructor",{value:x,configurable:!0}),x.displayName=p(b,l,"GeneratorFunction"),r.isGeneratorFunction=function(e){var t="function"==typeof e&&e.constructor;return!!t&&(t===x||"GeneratorFunction"===(t.displayName||t.name))},r.mark=function(e){return Object.setPrototypeOf?Object.setPrototypeOf(e,b):(e.__proto__=b,p(e,l,"GeneratorFunction")),e.prototype=Object.create(T),e},r.awrap=function(e){return{__await:e}},I(S.prototype),p(S.prototype,u,(function(){return this})),r.AsyncIterator=S,r.async=function(e,t,n,o,c){void 0===c&&(c=Promise);var i=new S(f(e,t,n,o),c);return r.isGeneratorFunction(t)?i:i.next().then((function(e){return e.done?e.value:i.next()}))},I(T),p(T,l,"Generator"),p(T,s,(function(){return this})),p(T,"toString",(function(){return"[object Generator]"})),r.keys=function(e){var t=Object(e),n=[];for(var r in t)n.push(r);return n.reverse(),function e(){for(;n.length;){var r=n.pop();if(r in t)return e.value=r,e.done=!1,e}return e.done=!0,e}},r.values=O,j.prototype={constructor:j,reset:function(e){if(this.prev=0,this.next=0,this.sent=this._sent=n,this.done=!1,this.delegate=null,this.method="next",this.arg=n,this.tryEntries.forEach(N),!e)for(var t in this)"t"===t.charAt(0)&&c.call(this,t)&&!isNaN(+t.slice(1))&&(this[t]=n)},stop:function(){this.done=!0;var e=this.tryEntries[0].completion;if("throw"===e.type)throw e.arg;return this.rval},dispatchException:function(e){if(this.done)throw e;var t=this;function r(r,o){return a.type="throw",a.arg=e,t.next=r,o&&(t.method="next",t.arg=n),!!o}for(var o=this.tryEntries.length-1;o>=0;--o){var i=this.tryEntries[o],a=i.completion;if("root"===i.tryLoc)return r("end");if(i.tryLoc<=this.prev){var s=c.call(i,"catchLoc"),u=c.call(i,"finallyLoc");if(s&&u){if(this.prev<i.catchLoc)return r(i.catchLoc,!0);if(this.prev<i.finallyLoc)return r(i.finallyLoc)}else if(s){if(this.prev<i.catchLoc)return r(i.catchLoc,!0)}else{if(!u)throw Error("try statement without catch or finally");if(this.prev<i.finallyLoc)return r(i.finallyLoc)}}}},abrupt:function(e,t){for(var n=this.tryEntries.length-1;n>=0;--n){var r=this.tryEntries[n];if(r.tryLoc<=this.prev&&c.call(r,"finallyLoc")&&this.prev<r.finallyLoc){var o=r;break}}o&&("break"===e||"continue"===e)&&o.tryLoc<=t&&t<=o.finallyLoc&&(o=null);var i=o?o.completion:{};return i.type=e,i.arg=t,o?(this.method="next",this.next=o.finallyLoc,y):this.complete(i)},complete:function(e,t){if("throw"===e.type)throw e.arg;return"break"===e.type||"continue"===e.type?this.next=e.arg:"return"===e.type?(this.rval=this.arg=e.arg,this.method="return",this.next="end"):"normal"===e.type&&t&&(this.next=t),y},finish:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var n=this.tryEntries[t];if(n.finallyLoc===e)return this.complete(n.completion,n.afterLoc),N(n),y}},catch:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var n=this.tryEntries[t];if(n.tryLoc===e){var r=n.completion;if("throw"===r.type){var o=r.arg;N(n)}return o}}throw Error("illegal catch attempt")},delegateYield:function(e,t,r){return this.delegate={iterator:O(e),resultName:t,nextLoc:r},"next"===this.method&&(this.arg=n),y}},r}function n(e,t,n,r,o,c,i){try{var a=e[c](i),s=a.value}catch(e){return void n(e)}a.done?t(s):Promise.resolve(s).then(r,o)}function r(e){return function(){var t=this,r=arguments;return new Promise((function(o,c){var i=e.apply(t,r);function a(e){n(i,o,c,a,s,"next",e)}function s(e){n(i,o,c,a,s,"throw",e)}a(void 0)}))}}var o=!1,c=null,i="Table has been generated in document",a="S.O.W. content generated in document",s="No table is selected in the document",u="Requested changes have been made in the document",l='<div class="bot-wrapper">\n  <img width=20 height=20 src="../../assets/copilot.png"/> NoviWord\n</div>\n<div class="message bot">',p=function(e){return'\n  <div class="bot-wrapper">\n    <img width="20" height="20" src="../../assets/copilot.png"/> NoviWord\n  </div>\n  <div class="message bot">'.concat(e,"</div>\n")};function f(e,t,n,r){return d.apply(this,arguments)}function d(){return(d=r(t().mark((function e(n,c,f,d){var m,v;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if(m=document.getElementById("chatWindow"),console.log("displayfunction called"),!(c&&c.attachments&&c.attachments.length>0)){e.next=6;break}c.attachments.forEach((function(e){e.content&&e.content.buttons&&e.content.buttons.length>0&&e.content.buttons.forEach((function(t){if("signin"===t.type){var n=document.createElement("button");n.innerText=t.title||"Sign In",n.classList.add("ms-Button","ms-Button--primary"),n.onclick=function(){window.open(t.value,"_blank")},m.innerHTML+='<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviPilot</div><div class="message bot">'.concat(e.content.text,"</div>"),m.appendChild(n)}}))})),e.next=32;break;case 6:if("bot"!==f){e.next=31;break}if("Generate"!==c.speak){e.next=13;break}h(c.text),m.innerHTML+="".concat(l).concat(a,"</div>"),o&&N((function(){W(a)})),e.next=29;break;case 13:if("Table"!==c.speak){e.next=19;break}g(c.text,"end"),m.innerHTML+="".concat(l).concat(i,"</div>"),o&&N((function(){W(i)})),e.next=29;break;case 19:if("TableReplace"!==c.speak){e.next=28;break}return v=!1,e.next=23,g(c.text,"replace");case 23:v=e.sent,console.log(v),v?(m.innerHTML+="".concat(l).concat(i,"</div>"),o&&N((function(){W(i)}))):(m.innerHTML+="".concat(l).concat(s,"</div>"),o&&N((function(){W(s)}))),e.next=29;break;case 28:"Replace"===c.speak?(splitText=c.text,textArray=splitText.split("|"),b(textArray[0],textArray[1]),m.innerHTML+="".concat(l,"Replaced ").concat(textArray[0]," with ").concat(textArray[1]," </div>"),o&&N((function(){W("Replaced ".concat(textArray[0]," with ").concat(textArray[1]))}))):"Selected"===c.speak?"Table"===c.text?(console.log("fetching selected table"),T(d)):(console.log("fetching selected data"),L(d)):"paragraph"===c.speak?(S(c.text),m.innerHTML+="".concat(l).concat(u,"</div>"),o&&N((function(){W(u)}))):"interim"===c.speak?(m.innerHTML+=p(c.text),o&&N(r(t().mark((function e(){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.next=2,W(c.text);case 2:o=!0;case 3:case"end":return e.stop()}}),e)}))))):"interimFinal"===c.speak?(m.innerHTML+=p(c.text),o&&N(r(t().mark((function e(){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:W(c.text);case 1:case"end":return e.stop()}}),e)}))))):c.text&&(m.innerHTML+=p(c.text),document.getElementById("insertButton").style.display="block",o&&(console.log("speaking bot message"),N((function(){W(c.text)}))));case 29:e.next=32;break;case 31:n&&(document.getElementById("insertButton").style.display="none",m.innerHTML+='<div class="user-wrapper">You</div><div class="message user">'.concat(n,"</div>"));case 32:x(),document.getElementById("userInput").value="";case 34:case"end":return e.stop()}}),e)})))).apply(this,arguments)}function h(e){return m.apply(this,arguments)}function m(){return m=r(t().mark((function e(n){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.next=2,Word.run(function(){var e=r(t().mark((function e(r){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return r.document.body.insertHtml(n,Word.InsertLocation.end),e.next=4,r.sync();case 4:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 2:case"end":return e.stop()}}),e)}))),m.apply(this,arguments)}function g(e,t){return v.apply(this,arguments)}function v(){return v=r(t().mark((function e(n,o){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if("end"!==o){e.next=7;break}return console.log("end of doc table"),e.next=4,Word.run(function(){var e=r(t().mark((function e(r){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return r.document.body.insertHtml(n,Word.InsertLocation.end),e.next=4,r.sync();case 4:return e.abrupt("return",!0);case 5:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 4:case 9:return e.abrupt("return",e.sent);case 7:return e.next=9,Word.run(function(){var e=r(t().mark((function e(r){var o,c,i,a;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return(o=r.document.getSelection()).load("parentTable"),e.next=4,r.sync();case 4:if(o.parentTable){e.next=7;break}return console.log("❌ No table selected."),e.abrupt("return",!1);case 7:return c=o.parentTable,(i=c.getRange(Word.RangeLocation.entire)).load(),e.next=12,r.sync();case 12:return console.log("Table found. Deleting..."),(a=i.insertText(" ",Word.InsertLocation.before)).load("text, address"),e.next=17,r.sync();case 17:return c.delete(),e.next=20,r.sync();case 20:return console.log("Table deleted."),console.log("Inserting new table..."),a.insertHtml(n,Word.InsertLocation.replace),e.next=25,r.sync();case 25:return console.log("New table inserted."),e.abrupt("return",!0);case 27:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 10:case"end":return e.stop()}}),e)}))),v.apply(this,arguments)}Office.onReady(function(){var e=r(t().mark((function e(n){var i;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return chatWindow.innerHTML+="".concat(l).concat("Hi! I'm NoviPilot, your Word assistant bot. I can help you create documents, modify content, and insert useful information seamlessly. How can I assist you today?","</div>"),e.next=3,y();case 3:i=e.sent,n.host===Office.HostType.Word&&(document.getElementById("sendButton").onclick=r(t().mark((function e(){var n;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if(!(n=document.getElementById("userInput").value)){e.next=5;break}return f(n,"","User",i),e.next=5,w(i,n);case 5:case"end":return e.stop()}}),e)}))),document.getElementById("userInput").addEventListener("keydown",function(){var e=r(t().mark((function e(n){var r;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if("Enter"!==n.key){e.next=7;break}if(n.preventDefault(),!(r=document.getElementById("userInput").value)){e.next=7;break}return f(r,"","User",i),e.next=7,w(i,r);case 7:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}()),document.getElementById("insertButton").onclick=r(t().mark((function e(){var n;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if(!(n=document.getElementById("chatWindow").lastChild?document.getElementById("chatWindow").lastChild.innerText:"")){e.next=4;break}return e.next=4,h(n);case 4:case"end":return e.stop()}}),e)}))),document.getElementById("startSpeechButton").addEventListener("click",(function(){if(o){mic.classList.toggle("recording"),o=!1;try{c.close()}catch(e){console.log("popup closed")}}else mic.classList.toggle("recording"),c=window.open("speech.html","SpeechRecognition","width=1,height=1"),o=!0,window.addEventListener("message",function(){var e=r(t().mark((function e(r){var a;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if(r.origin===window.location.origin){e.next=2;break}return e.abrupt("return");case 2:if(transcript=r.data,console.log(transcript),"NoSpeech"!==transcript){e.next=11;break}c.close(),o=!1,mic.classList.toggle("recording"),window.removeEventListener("message",n),e.next=23;break;case 11:if(document.getElementById("userInput").value=transcript,!(a=document.getElementById("userInput").value)){e.next=20;break}return document.getElementById("userInput").value="",f(a,"","User"),document.getElementById("startSpeechButton").disabled=!0,e.next=20,w(i,a);case 20:c.close(),mic.classList.toggle("recording"),window.removeEventListener("message",n);case 23:case"end":return e.stop()}}),e)})));function n(t){return e.apply(this,arguments)}return n}(),{once:!0})})));case 5:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());var y=function(){var e=r(t().mark((function e(){var n,r,o;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,fetch("https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview");case 3:return n=e.sent,e.next=6,n.json();case 6:if(r=e.sent,(o=new window.DirectLine.DirectLine({token:r.token}))&&o.activity$){e.next=10;break}throw new Error("DirectLine instance failed to initialize");case 10:return o.postActivity({from:{id:"10",name:"User"},type:"message",text:"Hi"}).subscribe((function(e){return console.log("Message sent with ID:",e)}),(function(e){return console.error("Error sending message:",e)})),o.activity$.subscribe((function(e){if(console.log("Testing activity on send: ",e),console.log("Role",e.from.role),"message"===e.type&&"10"!==e.from.id&&!e.recipient){console.log("Bot Response: ",e.text);try{document.getElementById("loader-container").remove(),document.getElementById("loader").remove()}catch(e){console.log("could not find the element")}f(!1,e,e.from.role,o)}})),e.abrupt("return",o);case 15:e.prev=15,e.t0=e.catch(0),console.error("Error initializing DirectLine:",e.t0);case 18:case"end":return e.stop()}}),e,null,[[0,15]])})));return function(){return e.apply(this,arguments)}}(),w=function(){var e=r(t().mark((function e(n,r){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:console.log("Calling loading function"),j(),console.log("User:",r),n.postActivity({from:{id:"10",name:"User"},type:"message",text:r}).subscribe((function(e){return console.log("Message sent with ID:",e)}),(function(e){return console.error("Error sending message:",e)}));case 4:case"end":return e.stop()}}),e)})));return function(t,n){return e.apply(this,arguments)}}();function x(){var e=document.getElementById("chatWindow");setTimeout((function(){e.scrollTop=e.scrollHeight}),100)}function b(e,t){return k.apply(this,arguments)}function k(){return k=r(t().mark((function e(n,o){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.next=2,Word.run(function(){var e=r(t().mark((function e(r){var c,i,a;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return(c=r.document.body.search(n,{matchCase:!1})).load("items"),e.next=4,r.sync();case 4:for(console.log("Results found:",c.items.length),i=c.items.length-1;i>=0;i--)a=c.items[i],console.log("Replacing:",a.text),a.insertText(o,Word.InsertLocation.replace);return e.next=8,r.sync();case 8:console.log("ll instances replaced successfully.");case 9:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 2:case"end":return e.stop()}}),e)}))),k.apply(this,arguments)}function L(e){return E.apply(this,arguments)}function E(){return E=r(t().mark((function e(n){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.next=2,Word.run(function(){var e=r(t().mark((function e(r){var o;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return(o=r.document.getSelection()).load("text"),e.next=4,r.sync();case 4:return SelText=o.text,e.next=7,w(n,SelText);case 7:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 2:case"end":return e.stop()}}),e)}))),E.apply(this,arguments)}function T(e){return I.apply(this,arguments)}function I(){return I=r(t().mark((function e(n){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.next=2,Word.run(function(){var e=r(t().mark((function e(r){var o,c,i,a;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return(o=r.document.getSelection()).load("parentTable"),e.next=4,r.sync();case 4:if(!o.parentTable){e.next=17;break}return(c=o.parentTable).load("values"),e.next=9,r.sync();case 9:return i=c.values,a="",i.forEach((function(e){a+=e.join(" | ")+"\n"})),console.log(a),e.next=15,w(n,a);case 15:e.next=18;break;case 17:console.log("No table selected by user.");case 18:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 2:case"end":return e.stop()}}),e)}))),I.apply(this,arguments)}function S(e){return B.apply(this,arguments)}function B(){return B=r(t().mark((function e(n){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.next=2,Word.run(function(){var e=r(t().mark((function e(r){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return r.document.getSelection().insertText(n,Word.InsertLocation.replace),e.next=4,r.sync();case 4:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 2:case"end":return e.stop()}}),e)}))),B.apply(this,arguments)}function W(e){return H.apply(this,arguments)}function H(){return(H=r(t().mark((function e(n){var r,c,i;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return console.log("Testing Text to Speech"),r=window.speechSynthesis.getVoices(),console.log("Voices:",r),c=r.find((function(e){return e.name.includes("Female")||e.name.includes("Google UK English Female")||e.name.includes("Microsoft Zira")||e.name.includes("Samantha")})),console.log("Set voice:",c),i=new SpeechSynthesisUtterance(n),c?i.voice=c:console.warn("Female voice not found. Using default voice."),e.abrupt("return",new Promise((function(e){i.onend=function(){console.log("Speech has finished."),o=!1,document.getElementById("startSpeechButton").disabled=!1,console.log("mic enabled"),e(!0)},i.onerror=function(t){console.error("Speech error:",t.error),o=!1,document.getElementById("startSpeechButton").disabled=!1,console.log("mic enabled"),e(!1)},window.speechSynthesis.speak(i)})));case 8:case"end":return e.stop()}}),e)})))).apply(this,arguments)}function N(e){window.speechSynthesis.getVoices().length>0?e():window.speechSynthesis.onvoiceschanged=e}function j(){var e=document.getElementById("chatWindow"),t=[".","..","..."],n=1;e.innerHTML+='<div id="loader-container" class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviPilot</div><div id = "loader" class="message bot">.</div>';var r=document.getElementById("loader");setInterval((function(){r.innerText=t[n],n=(n+1)%t.length}),500)}}(),function(){"use strict";new URL(n(58394),n.b),new URL(n(84400),n.b),new URL(n(60947),n.b)}()}();
//# sourceMappingURL=taskpane.js.map
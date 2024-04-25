/*! For license information please see taskpane.js.LICENSE.txt */
!function(){"use strict";var e,t,n,r,a={27091:function(e){e.exports=function(e,t){return t||(t={}),e?(e=String(e.__esModule?e.default:e),t.hash&&(e+=t.hash),t.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(e)?'"'.concat(e,'"'):e):e}},96384:function(e,t,n){e.exports=n.p+"assets/space-station.png"},60806:function(e,t,n){e.exports=n.p+"a2de33a19b17be3c06e2.css"}},o={};function c(e){var t=o[e];if(void 0!==t)return t.exports;var n=o[e]={exports:{}};return a[e](n,n.exports,c),n.exports}c.m=a,c.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return c.d(t,{a:t}),t},c.d=function(e,t){for(var n in t)c.o(t,n)&&!c.o(e,n)&&Object.defineProperty(e,n,{enumerable:!0,get:t[n]})},c.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),c.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;c.g.importScripts&&(e=c.g.location+"");var t=c.g.document;if(!e&&t&&(t.currentScript&&(e=t.currentScript.src),!e)){var n=t.getElementsByTagName("script");if(n.length)for(var r=n.length-1;r>-1&&!e;)e=n[r--].src}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),c.p=e}(),c.b=document.baseURI||self.location.href,function(){class e extends Error{}function t(t,n){if("string"!=typeof t)throw new e("Invalid token specified: must be a string");n||(n={});const r=!0===n.header?0:1,a=t.split(".")[r];if("string"!=typeof a)throw new e(`Invalid token specified: missing part #${r+1}`);let o;try{o=function(e){let t=e.replace(/-/g,"+").replace(/_/g,"/");switch(t.length%4){case 0:break;case 2:t+="==";break;case 3:t+="=";break;default:throw new Error("base64 string is not of the correct length")}try{return function(e){return decodeURIComponent(atob(e).replace(/(.)/g,((e,t)=>{let n=t.charCodeAt(0).toString(16).toUpperCase();return n.length<2&&(n="0"+n),"%"+n})))}(t)}catch(e){return atob(t)}}(a)}catch(t){throw new e(`Invalid token specified: invalid base64 for part #${r+1} (${t.message})`)}try{return JSON.parse(o)}catch(t){throw new e(`Invalid token specified: invalid json for part #${r+1} (${t.message})`)}}function n(e){return n="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(e){return typeof e}:function(e){return e&&"function"==typeof Symbol&&e.constructor===Symbol&&e!==Symbol.prototype?"symbol":typeof e},n(e)}function r(){r=function(){return t};var e,t={},a=Object.prototype,o=a.hasOwnProperty,c=Object.defineProperty||function(e,t,n){e[t]=n.value},i="function"==typeof Symbol?Symbol:{},s=i.iterator||"@@iterator",l=i.asyncIterator||"@@asyncIterator",u=i.toStringTag||"@@toStringTag";function f(e,t,n){return Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}),e[t]}try{f({},"")}catch(e){f=function(e,t,n){return e[t]=n}}function d(e,t,n,r){var a=t&&t.prototype instanceof b?t:b,o=Object.create(a.prototype),i=new S(r||[]);return c(o,"_invoke",{value:j(e,n,i)}),o}function h(e,t,n){try{return{type:"normal",arg:e.call(t,n)}}catch(e){return{type:"throw",arg:e}}}t.wrap=d;var p="suspendedStart",y="suspendedYield",m="executing",g="completed",x={};function b(){}function v(){}function w(){}var k={};f(k,s,(function(){return this}));var E=Object.getPrototypeOf,C=E&&E(E(A([])));C&&C!==a&&o.call(C,s)&&(k=C);var O=w.prototype=b.prototype=Object.create(k);function B(e){["next","throw","return"].forEach((function(t){f(e,t,(function(e){return this._invoke(t,e)}))}))}function I(e,t){function r(a,c,i,s){var l=h(e[a],e,c);if("throw"!==l.type){var u=l.arg,f=u.value;return f&&"object"==n(f)&&o.call(f,"__await")?t.resolve(f.__await).then((function(e){r("next",e,i,s)}),(function(e){r("throw",e,i,s)})):t.resolve(f).then((function(e){u.value=e,i(u)}),(function(e){return r("throw",e,i,s)}))}s(l.arg)}var a;c(this,"_invoke",{value:function(e,n){function o(){return new t((function(t,a){r(e,n,t,a)}))}return a=a?a.then(o,o):o()}})}function j(t,n,r){var a=p;return function(o,c){if(a===m)throw new Error("Generator is already running");if(a===g){if("throw"===o)throw c;return{value:e,done:!0}}for(r.method=o,r.arg=c;;){var i=r.delegate;if(i){var s=L(i,r);if(s){if(s===x)continue;return s}}if("next"===r.method)r.sent=r._sent=r.arg;else if("throw"===r.method){if(a===p)throw a=g,r.arg;r.dispatchException(r.arg)}else"return"===r.method&&r.abrupt("return",r.arg);a=m;var l=h(t,n,r);if("normal"===l.type){if(a=r.done?g:y,l.arg===x)continue;return{value:l.arg,done:r.done}}"throw"===l.type&&(a=g,r.method="throw",r.arg=l.arg)}}}function L(t,n){var r=n.method,a=t.iterator[r];if(a===e)return n.delegate=null,"throw"===r&&t.iterator.return&&(n.method="return",n.arg=e,L(t,n),"throw"===n.method)||"return"!==r&&(n.method="throw",n.arg=new TypeError("The iterator does not provide a '"+r+"' method")),x;var o=h(a,t.iterator,n.arg);if("throw"===o.type)return n.method="throw",n.arg=o.arg,n.delegate=null,x;var c=o.arg;return c?c.done?(n[t.resultName]=c.value,n.next=t.nextLoc,"return"!==n.method&&(n.method="next",n.arg=e),n.delegate=null,x):c:(n.method="throw",n.arg=new TypeError("iterator result is not an object"),n.delegate=null,x)}function N(e){var t={tryLoc:e[0]};1 in e&&(t.catchLoc=e[1]),2 in e&&(t.finallyLoc=e[2],t.afterLoc=e[3]),this.tryEntries.push(t)}function R(e){var t=e.completion||{};t.type="normal",delete t.arg,e.completion=t}function S(e){this.tryEntries=[{tryLoc:"root"}],e.forEach(N,this),this.reset(!0)}function A(t){if(t||""===t){var r=t[s];if(r)return r.call(t);if("function"==typeof t.next)return t;if(!isNaN(t.length)){var a=-1,c=function n(){for(;++a<t.length;)if(o.call(t,a))return n.value=t[a],n.done=!1,n;return n.value=e,n.done=!0,n};return c.next=c}}throw new TypeError(n(t)+" is not iterable")}return v.prototype=w,c(O,"constructor",{value:w,configurable:!0}),c(w,"constructor",{value:v,configurable:!0}),v.displayName=f(w,u,"GeneratorFunction"),t.isGeneratorFunction=function(e){var t="function"==typeof e&&e.constructor;return!!t&&(t===v||"GeneratorFunction"===(t.displayName||t.name))},t.mark=function(e){return Object.setPrototypeOf?Object.setPrototypeOf(e,w):(e.__proto__=w,f(e,u,"GeneratorFunction")),e.prototype=Object.create(O),e},t.awrap=function(e){return{__await:e}},B(I.prototype),f(I.prototype,l,(function(){return this})),t.AsyncIterator=I,t.async=function(e,n,r,a,o){void 0===o&&(o=Promise);var c=new I(d(e,n,r,a),o);return t.isGeneratorFunction(n)?c:c.next().then((function(e){return e.done?e.value:c.next()}))},B(O),f(O,u,"Generator"),f(O,s,(function(){return this})),f(O,"toString",(function(){return"[object Generator]"})),t.keys=function(e){var t=Object(e),n=[];for(var r in t)n.push(r);return n.reverse(),function e(){for(;n.length;){var r=n.pop();if(r in t)return e.value=r,e.done=!1,e}return e.done=!0,e}},t.values=A,S.prototype={constructor:S,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=e,this.done=!1,this.delegate=null,this.method="next",this.arg=e,this.tryEntries.forEach(R),!t)for(var n in this)"t"===n.charAt(0)&&o.call(this,n)&&!isNaN(+n.slice(1))&&(this[n]=e)},stop:function(){this.done=!0;var e=this.tryEntries[0].completion;if("throw"===e.type)throw e.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var n=this;function r(r,a){return i.type="throw",i.arg=t,n.next=r,a&&(n.method="next",n.arg=e),!!a}for(var a=this.tryEntries.length-1;a>=0;--a){var c=this.tryEntries[a],i=c.completion;if("root"===c.tryLoc)return r("end");if(c.tryLoc<=this.prev){var s=o.call(c,"catchLoc"),l=o.call(c,"finallyLoc");if(s&&l){if(this.prev<c.catchLoc)return r(c.catchLoc,!0);if(this.prev<c.finallyLoc)return r(c.finallyLoc)}else if(s){if(this.prev<c.catchLoc)return r(c.catchLoc,!0)}else{if(!l)throw new Error("try statement without catch or finally");if(this.prev<c.finallyLoc)return r(c.finallyLoc)}}}},abrupt:function(e,t){for(var n=this.tryEntries.length-1;n>=0;--n){var r=this.tryEntries[n];if(r.tryLoc<=this.prev&&o.call(r,"finallyLoc")&&this.prev<r.finallyLoc){var a=r;break}}a&&("break"===e||"continue"===e)&&a.tryLoc<=t&&t<=a.finallyLoc&&(a=null);var c=a?a.completion:{};return c.type=e,c.arg=t,a?(this.method="next",this.next=a.finallyLoc,x):this.complete(c)},complete:function(e,t){if("throw"===e.type)throw e.arg;return"break"===e.type||"continue"===e.type?this.next=e.arg:"return"===e.type?(this.rval=this.arg=e.arg,this.method="return",this.next="end"):"normal"===e.type&&t&&(this.next=t),x},finish:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var n=this.tryEntries[t];if(n.finallyLoc===e)return this.complete(n.completion,n.afterLoc),R(n),x}},catch:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var n=this.tryEntries[t];if(n.tryLoc===e){var r=n.completion;if("throw"===r.type){var a=r.arg;R(n)}return a}}throw new Error("illegal catch attempt")},delegateYield:function(t,n,r){return this.delegate={iterator:A(t),resultName:n,nextLoc:r},"next"===this.method&&(this.arg=e),x}},t}function a(e,t,n,r,a,o,c){try{var i=e[o](c),s=i.value}catch(e){return void n(e)}i.done?t(s):Promise.resolve(s).then(r,a)}function o(e){return function(){var t=this,n=arguments;return new Promise((function(r,o){var c=e.apply(t,n);function i(e){a(c,r,o,i,s,"next",e)}function s(e){a(c,r,o,i,s,"throw",e)}i(void 0)}))}}e.prototype.name="InvalidTokenError",(new Date).getMonth();var c=["#C1B","#C1A","#C2B","#C2A","#C3B","#C3A","#C4B","#C4A","#C5B","#C5A","#C6B","#C6A","#C7B","#C7A","#C8B","#C8A","#C9B","#C9A","#C10B","#C10A","#C11B","#C11A","#C12B","#C12A"];function i(){return s.apply(this,arguments)}function s(){return s=o(r().mark((function e(){return r().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(r().mark((function e(t){var n,a,o,i,s,l,u,f,d,h,p,y,m,g,x,b,v,w,k,E,C,O,B,I,j,L,N,R,S,A,D,P,T,_,M,G,$,U,F,V,X,z,J,W,Y,H,Q,q,K,Z;return r().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if("nil"!=document.getElementById("budgetList").value){e.next=5;break}document.getElementById("load-status").textContent="Please select a user and log in",document.getElementById("load-bar").style.width="0%",e.next=197;break;case 5:if(""!=document.getElementById("fname").value){e.next=10;break}document.getElementById("load-status").textContent="Please enter the URL for the data",document.getElementById("load-bar").style.width="0%",e.next=197;break;case 10:return e.next=12,OfficeRuntime.auth.getAccessToken({allowSignInPrompt:!0,allowConsentPrompt:!0});case 12:return e.sent,n=t.workbook.worksheets.getActiveWorksheet(),a=[],o=[],i=[],e.next=19,t.sync();case 19:document.getElementById("load-status").textContent="loading...",document.getElementById("load-bar").style.width="0%",s=3.5,l=3.5,u=0,f=0,d=n.getRange(),h=!1,p=!1,y=!1,m=0,g="",x=[-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1],b=document.getElementById("budgetList").value,v=document.getElementById("budgetList").options[document.getElementById("budgetList").selectedIndex].text,w=document.getElementById("fname").value,k=0;case 36:if(!(k<24)){e.next=58;break}m=0,h=!1,p=!1,g=c[k];case 41:if(h||p||y){e.next=51;break}return E=d.getColumn(m),C=E.findOrNullObject(g,{completeMatch:!0,matchCase:!0,searchDirection:Excel.SearchDirection.forward}),e.next=46,t.sync();case 46:C.isNullObject||(h=!0,x[k]=m,m+1>u&&(u=m+1)),++m>=500&&(p=!0,y=!0),e.next=41;break;case 51:return document.getElementById("load-bar").style.width=s+"%",e.next=54,t.sync();case 54:s+=l;case 55:k++,e.next=36;break;case 58:if(y){e.next=195;break}m=0,h=!1,p=!1;case 62:if(h||p){e.next=72;break}return O=d.getRow(m),B=O.findOrNullObject("#RX",{completeMatch:!0,matchCase:!0,searchDirection:Excel.SearchDirection.forward}),e.next=67,t.sync();case 67:B.isNullObject||(h=!0,f=m+1),++m>=500&&(p=!0,f=m+1),e.next=62;break;case 72:i=Array(f).fill().map((function(){return Array(u).fill(!1)})),m=0,p=!1;case 75:if(p||!(m<u)){e.next=90;break}return I=d.getColumn(m),j=I.findOrNullObject("#N",{completeMatch:!0,matchCase:!0,searchDirection:Excel.SearchDirection.forward}),e.next=80,t.sync();case 80:if(j.isNullObject){e.next=87;break}return(L=d.getCell(0,m).getResizedRange(f,0)).load("text"),e.next=85,t.sync();case 85:a=L.text,p=!0;case 87:m++,e.next=75;break;case 90:return document.getElementById("load-bar").style.width=s+"%",e.next=93,t.sync();case 93:s+=l,m=0,p=!1;case 96:if(p||!(m<u)){e.next=111;break}return N=d.getColumn(m),R=N.findOrNullObject("#RC",{completeMatch:!0,matchCase:!0,searchDirection:Excel.SearchDirection.forward}),e.next=101,t.sync();case 101:if(R.isNullObject){e.next=108;break}return(S=d.getCell(0,m).getResizedRange(f,0)).load("text"),e.next=106,t.sync();case 106:o=S.text,p=!0;case 108:m++,e.next=96;break;case 111:return A="",D=!1,e.next=115,fetch(w,{method:"POST",headers:{"Content-Type":"application/json"},body:'{ "param": "'+b+'", "mode": "BVA" }'}).then((function(e){return e.json()})).then((function(e){return JSON.stringify(e)})).then((function(e){A=e})).catch((function(e){D=!0}));case 115:if(D){e.next=191;break}return document.getElementById("load-bar").style.width=s+"%",e.next=119,t.sync();case 119:if(s+=l,P=JSON.parse(A),"[]"==A){e.next=187;break}return document.getElementById("load-bar").style.width=s+"%",e.next=125,t.sync();case 125:for(s+=l,T=null,_="",M="",G=[],$=-1,U=0,F=!1,V="",X=0,z=0;z<P.length;z++)if(T=P[z],_=T.rpg,M=T.Date,3==(G=M.split("/")).length&&!G[1].isNaN){for($=x[2*(parseInt(G[1])-1)],p=!1,a[X][0]!=_&&(X=0);!p&&X<f;)a[X][0]==_?(U=0,T.hasOwnProperty("BudgetValue")&&(U=T.BudgetValue),d.getCell(X,$).values=[[U]],i[X][$]=!0,U=0,T.hasOwnProperty("ActualValue")&&(U=T.ActualValue),$=x[2*(parseInt(G[1])-1)+1],d.getCell(X,$).values=[[U]],i[X][$]=!0,p=!0):X++;p||(F=!0,""==V&&(V=_))}J=0;case 137:if(!(J<f)){e.next=154;break}if("#R"!=o[J][0]&&"#RX"!=o[J][0]){e.next=151;break}W=0;case 140:if(!(W<24)){e.next=151;break}if(-1==x[W]){e.next=148;break}if(i[J][x[W]]){e.next=148;break}return(Y=d.getCell(J,x[W])).load("valueTypes"),e.next=147,t.sync();case 147:Y.valueTypes[0][0]!=Excel.RangeValueType.empty&&(Y.values=[[""]]);case 148:W++,e.next=140;break;case 151:J++,e.next=137;break;case 154:m=0,p=!1;case 156:if(p||!(m<u)){e.next=184;break}return H=d.getColumn(m),Q=H.findOrNullObject("#BC",{completeMatch:!0,matchCase:!0,searchDirection:Excel.SearchDirection.forward}),e.next=161,t.sync();case 161:if(Q.isNullObject){e.next=181;break}q=m,m=0;case 164:if(p||!(m<500)){e.next=181;break}return K=d.getRow(m),Z=K.findOrNullObject("#BR",{completeMatch:!0,matchCase:!0,searchDirection:Excel.SearchDirection.forward}),e.next=169,t.sync();case 169:if(Z.isNullObject){e.next=178;break}return p=!0,dataRange=d.getCell(m,q),dataRange.clear(Excel.ClearApplyTo.contents),e.next=175,t.sync();case 175:return dataRange.values=[[v]],e.next=178,t.sync();case 178:m++,e.next=164;break;case 181:m++,e.next=156;break;case 184:F?(document.getElementById("load-status").textContent="Warning. RPG code "+V+" missing",document.getElementById("load-bar").style.width="100%"):(document.getElementById("load-status").textContent="Loaded successfully.",document.getElementById("load-bar").style.width="100%"),e.next=189;break;case 187:document.getElementById("load-status").textContent="Load failed. No data in resource.",document.getElementById("load-bar").style.width="0%";case 189:e.next=193;break;case 191:document.getElementById("load-status").textContent="Load failed. Resource at URL is unreachable.",document.getElementById("load-bar").style.width="0%";case 193:e.next=197;break;case 195:document.getElementById("load-status").textContent="Load failed. Missing column markers.",document.getElementById("load-bar").style.width="0%";case 197:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=8;break;case 5:e.prev=5,e.t0=e.catch(0),console.error(e.t0);case 8:case"end":return e.stop()}}),e,null,[[0,5]])}))),s.apply(this,arguments)}function l(){return u.apply(this,arguments)}function u(){return u=o(r().mark((function e(){return r().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(r().mark((function e(t){var n,a,o,i,s,l,u,f,d,h,p,y,m,g,x,b,v,w,k,E,C,O,B,I,j;return r().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return n=t.workbook.worksheets.getActiveWorksheet(),a=[],e.next=4,t.sync();case 4:document.getElementById("load-status").textContent="",document.getElementById("load-bar").style.width="0%",document.getElementById("clear-status").textContent="clearing...",o=n.getRange(),i=!1,s=!1,l=0,u="",f=0,d=0,h=[-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1],p=0;case 16:if(!(p<24)){e.next=34;break}l=0,i=!1,s=!1,u=c[p];case 21:if(i||s){e.next=31;break}return y=o.getColumn(l),m=y.findOrNullObject(u,{completeMatch:!0,matchCase:!0,searchDirection:Excel.SearchDirection.forward}),e.next=26,t.sync();case 26:m.isNullObject||(i=!0,h[p]=l,l+1>f&&(f=l+1)),++l>=500&&(s=!0),e.next=21;break;case 31:p++,e.next=16;break;case 34:l=0,i=!1,s=!1;case 37:if(i||s){e.next=47;break}return g=o.getRow(l),x=g.findOrNullObject("#RX",{completeMatch:!0,matchCase:!0,searchDirection:Excel.SearchDirection.forward}),e.next=42,t.sync();case 42:x.isNullObject||(i=!0,d=l+1),++l>=500&&(s=!0,d=l+1),e.next=37;break;case 47:l=0,s=!1;case 49:if(s||!(l<f)){e.next=64;break}return b=o.getColumn(l),v=b.findOrNullObject("#RC",{completeMatch:!0,matchCase:!0,searchDirection:Excel.SearchDirection.forward}),e.next=54,t.sync();case 54:if(v.isNullObject){e.next=61;break}return(w=o.getCell(0,l).getResizedRange(d,0)).load("text"),e.next=59,t.sync();case 59:a=w.text,s=!0;case 61:l++,e.next=49;break;case 64:for(k=0;k<d;k++)if("#R"==a[k][0]||"#RX"==a[k][0])for(E=0;E<24;E++)-1!=h[E]&&(o.getCell(k,h[E]).values=[[""]]);l=0,s=!1;case 67:if(s||!(l<f)){e.next=92;break}return C=o.getColumn(l),O=C.findOrNullObject("#BC",{completeMatch:!0,matchCase:!0,searchDirection:Excel.SearchDirection.forward}),e.next=72,t.sync();case 72:if(O.isNullObject){e.next=89;break}B=l,l=0;case 75:if(s||!(l<d)){e.next=89;break}return I=o.getRow(l),j=I.findOrNullObject("#BR",{completeMatch:!0,matchCase:!0,searchDirection:Excel.SearchDirection.forward}),e.next=80,t.sync();case 80:if(j.isNullObject){e.next=86;break}return s=!0,dataRange=o.getCell(l,B),dataRange.clear(Excel.ClearApplyTo.contents),e.next=86,t.sync();case 86:l++,e.next=75;break;case 89:l++,e.next=67;break;case 92:document.getElementById("clear-status").textContent=" ";case 93:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=8;break;case 5:e.prev=5,e.t0=e.catch(0),console.error(e.t0);case 8:case"end":return e.stop()}}),e,null,[[0,5]])}))),u.apply(this,arguments)}function f(){return d.apply(this,arguments)}function d(){return(d=o(r().mark((function e(){var n,a;return r().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,console.log("Login Button Clicked"),e.next=4,Office.auth.getAccessToken({allowSignInPrompt:!0,allowConsentPrompt:!0,forMSGraphAccess:!0});case 4:n=e.sent,a=t(n),console.log("user",a),e.next=12;break;case 9:e.prev=9,e.t0=e.catch(0),console.error(e.t0);case 12:case"end":return e.stop()}}),e,null,[[0,9]])})))).apply(this,arguments)}Office.onReady((function(e){e.host===Office.HostType.Excel&&(document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("load").onclick=i,document.getElementById("clear").onclick=l,document.getElementById("login").onclick=f)}))}(),e=c(27091),t=c.n(e),n=new URL(c(60806),c.b),r=new URL(c(96384),c.b),t()(n),t()(r)}();
//# sourceMappingURL=taskpane.js.map
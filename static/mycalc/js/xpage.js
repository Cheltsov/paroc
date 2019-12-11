(window["webpackJsonp"] = window["webpackJsonp"] || []).push([["xpage"],{

/***/ "./src/ts/xpage/Element.ts":
/*!*********************************!*\
  !*** ./src/ts/xpage/Element.ts ***!
  \*********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

var __WEBPACK_AMD_DEFINE_ARRAY__, __WEBPACK_AMD_DEFINE_RESULT__;!(__WEBPACK_AMD_DEFINE_ARRAY__ = [__webpack_require__, exports, __webpack_require__(/*! ./core */ "./src/ts/xpage/core.ts")], __WEBPACK_AMD_DEFINE_RESULT__ = (function (require, exports, core_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    var Element = (function () {
        function Element(selector) {
            this._selector = "";
            this._length = 0;
            if (!selector)
                this.els = [];
            else if (typeof selector == "string") {
                this.els = core_1.default.elementsGetter(selector);
                this._selector = selector;
            }
            else if (selector instanceof HTMLElement)
                this.els = [selector];
            else if (selector instanceof NodeList)
                this.els = core_1.default.transformNodeListToArray(selector);
            else if (selector instanceof Array && (selector[0] instanceof HTMLElement || !selector.length))
                this.els = selector;
            else if (selector instanceof Element)
                this.els = selector.els;
            else
                throw Error("Invalid selector: " + selector);
        }
        Object.defineProperty(Element.prototype, "els", {
            get: function () {
                return this._els;
            },
            set: function (elements) {
                this._els = elements;
                this._length = elements.length || 0;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Element.prototype, "length", {
            get: function () {
                return this._length;
            },
            enumerable: true,
            configurable: true
        });
        Element.prototype.addElement = function (element) {
            if (typeof element == "string")
                this.els = this.els.concat(core_1.default.elementsGetter(element));
            else if (element instanceof Element)
                this.els = this.els.concat(element.els);
            else if (element instanceof HTMLElement)
                this.els = this.els.concat(element);
            else if (element instanceof NodeList)
                this.els = this.els.concat(core_1.default.transformNodeListToArray(element));
            else if (element instanceof Array && (element[0] instanceof HTMLElement || !element.length))
                this.els = this.els.concat(element);
            else
                throw Error("Invalid selector: " + element);
            return this;
        };
        Element.prototype.is = function (selector) {
            var element;
            if (typeof selector == "string")
                element = core_1.default.elementsGetter(selector);
            else if (selector instanceof HTMLElement)
                element = [selector];
            return this.els[0] == element[0];
        };
        Element.prototype.has = function (selector) {
            var searchElements;
            if (typeof selector == "string")
                searchElements = core_1.default.elementsGetter(selector);
            else if (selector instanceof HTMLElement)
                searchElements = [selector];
            else if (selector instanceof Element)
                searchElements = selector._els;
            else if (selector instanceof NodeList)
                searchElements = core_1.default.transformNodeListToArray(selector);
            else if (selector instanceof Array && (selector[0] instanceof HTMLElement || !selector.length))
                searchElements = selector;
            else
                throw Error("Invalid selector: " + selector);
            var isFinded = false;
            for (var _i = 0, _a = this.els; _i < _a.length; _i++) {
                var el = _a[_i];
                for (var _b = 0, searchElements_1 = searchElements; _b < searchElements_1.length; _b++) {
                    var target = searchElements_1[_b];
                    if (el.contains(target)) {
                        isFinded = true;
                        break;
                    }
                }
                if (isFinded)
                    return true;
            }
            return false;
        };
        Element.prototype.addClass = function (className) {
            core_1.default.each(this.els, function (el) {
                el.classList.add(className);
            });
            return this;
        };
        Element.prototype.removeClass = function (className) {
            core_1.default.each(this.els, function (el) {
                el.classList.remove(className);
            });
            return this;
        };
        Element.prototype.toggleClass = function (className, callback) {
            core_1.default.each(this.els, function (el) {
                if (el.classList.contains(className)) {
                    el.classList.remove(className);
                    if (callback)
                        callback(false);
                }
                else {
                    el.classList.add(className);
                    if (callback)
                        callback(true);
                }
            });
            return this;
        };
        Element.prototype.hasClass = function (targetClass) {
            for (var _i = 0, _a = this.els; _i < _a.length; _i++) {
                var el = _a[_i];
                if (el.classList.contains(targetClass))
                    return true;
            }
            return false;
        };
        Element.prototype.find = function (selector) {
            var searchingElements = new Array();
            core_1.default.each(this.els, function (el) {
                var findedElements = el.querySelectorAll(selector);
                if (!findedElements.length)
                    return;
                for (var _i = 0, _a = core_1.default.transformNodeListToArray(findedElements); _i < _a.length; _i++) {
                    var el_1 = _a[_i];
                    searchingElements.push(el_1);
                }
            });
            return new Element(searchingElements);
        };
        Element.prototype.closest = function (selector) {
            var searchingElements = new Element();
            core_1.default.each(this.els, function (el) {
                var findedElements = el.closest(selector);
                if (!findedElements)
                    return;
                searchingElements.addElement(findedElements);
            });
            return searchingElements;
        };
        Element.prototype.text = function (text) {
            if (text) {
                core_1.default.each(this.els, function (el) {
                    el.innerText = text;
                });
                return this;
            }
            else if (this.length > 1) {
                var textArray_1 = [];
                core_1.default.each(this.els, function (el) {
                    textArray_1.push(el.innerText);
                });
                return textArray_1;
            }
            else
                return this.els[0].innerText;
        };
        Element.prototype.get = function (index) {
            if (this.els[index])
                return new Element(this.els[index]);
            else
                return new Element();
        };
        Element.prototype.getHTMLElement = function (index) {
            return this.els[index];
        };
        Element.prototype.height = function (height) {
            if (!height)
                return parseInt(getComputedStyle(this.els[0]).height);
            core_1.default.each(this.els, function (el) {
                if (isNaN(height))
                    el.style.height = height;
                else
                    el.style.height = height + "px";
            });
            return this;
        };
        Element.prototype.map = function (callback) {
            return this.els.map(callback);
        };
        Element.prototype.attr = function (attrName, value) {
            if (value) {
                core_1.default.each(this.els, function (el) {
                    el.setAttribute(attrName, value);
                });
                return this;
            }
            return this.els[0].getAttribute(attrName);
        };
        Element.prototype.prev = function (selector) {
            var searchingElements = new Element();
            core_1.default.each(this.els, function (el) {
                var findedElements = el.previousElementSibling;
                if (!findedElements)
                    return;
                if (selector) {
                    if (findedElements.classList.contains(selector.replace(".", "")))
                        searchingElements.addElement(findedElements);
                }
                else
                    searchingElements.addElement(findedElements);
            });
            return searchingElements;
        };
        return Element;
    }());
    exports.default = Element;
}).apply(exports, __WEBPACK_AMD_DEFINE_ARRAY__),
				__WEBPACK_AMD_DEFINE_RESULT__ !== undefined && (module.exports = __WEBPACK_AMD_DEFINE_RESULT__));


/***/ }),

/***/ "./src/ts/xpage/EventListener.ts":
/*!***************************************!*\
  !*** ./src/ts/xpage/EventListener.ts ***!
  \***************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

var __WEBPACK_AMD_DEFINE_ARRAY__, __WEBPACK_AMD_DEFINE_RESULT__;var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
!(__WEBPACK_AMD_DEFINE_ARRAY__ = [__webpack_require__, exports, __webpack_require__(/*! ./core */ "./src/ts/xpage/core.ts"), __webpack_require__(/*! ./Element */ "./src/ts/xpage/Element.ts")], __WEBPACK_AMD_DEFINE_RESULT__ = (function (require, exports, core_1, Element_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    var EventListener = (function (_super) {
        __extends(EventListener, _super);
        function EventListener() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        EventListener.prototype.add = function (event, callback, options) {
            var self = this;
            if (this._selector)
                document.body.addEventListener(event, function (e) {
                    core_1.default.each(self._selector, function (el, i) {
                        var target = e.target;
                        if (el.contains(target)
                            || el == target)
                            callback(el, e, i);
                    });
                }, options);
            else
                core_1.default.each(this.els, function (el, i) {
                    document.body.addEventListener(event, function (e) {
                        var target = e.target;
                        if (el.contains(target)
                            || el == target)
                            callback(el, e, i);
                    }, options);
                });
            return this;
        };
        EventListener.prototype.trigger = function (event) {
            core_1.default.each(this.els, function (el) {
                var evt = document.createEvent("HTMLEvents");
                evt.initEvent(event, false, true);
                el.dispatchEvent(evt);
            });
            return this;
        };
        return EventListener;
    }(Element_1.default));
    exports.default = EventListener;
}).apply(exports, __WEBPACK_AMD_DEFINE_ARRAY__),
				__WEBPACK_AMD_DEFINE_RESULT__ !== undefined && (module.exports = __WEBPACK_AMD_DEFINE_RESULT__));


/***/ }),

/***/ "./src/ts/xpage/core.ts":
/*!******************************!*\
  !*** ./src/ts/xpage/core.ts ***!
  \******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

var __WEBPACK_AMD_DEFINE_ARRAY__, __WEBPACK_AMD_DEFINE_RESULT__;!(__WEBPACK_AMD_DEFINE_ARRAY__ = [__webpack_require__, exports], __WEBPACK_AMD_DEFINE_RESULT__ = (function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    var App = (function () {
        function App() {
        }
        App.getElements = function (selector) {
            var elements = document.querySelectorAll(selector);
            return elements.length ? elements : [];
        };
        App.getElement = function (selector) {
            var element = document.querySelector(selector);
            return element;
        };
        App.elementsGetter = function (selector) {
            return App.transformNodeListToArray(document.querySelectorAll(selector));
        };
        App.transformNodeListToArray = function (list) {
            try {
                return Array.prototype.slice.call(list);
            }
            catch (e) {
                throw Error(e);
                return [];
            }
        };
        App.wrap = function (selector, wrapper) {
            var localWrapper;
            if (typeof wrapper == "string")
                localWrapper = document.createElement(wrapper);
            else if (wrapper instanceof HTMLElement)
                localWrapper = wrapper;
            App.each(selector, function (el, i) {
                localWrapper.innerHTML = el.outerHTML;
                el.parentNode.replaceChild(localWrapper, el);
            });
        };
        App.each = function (elements, callback) {
            if (!callback) {
                console.error("Callback does not exists in yours 'each'");
                return;
            }
            if (typeof elements == "string")
                elements = App.transformNodeListToArray(App.getElements(elements));
            var i = 0;
            for (var _i = 0, elements_1 = elements; _i < elements_1.length; _i++) {
                var el = elements_1[_i];
                callback(el, i);
                i++;
            }
        };
        return App;
    }());
    exports.default = App;
}).apply(exports, __WEBPACK_AMD_DEFINE_ARRAY__),
				__WEBPACK_AMD_DEFINE_RESULT__ !== undefined && (module.exports = __WEBPACK_AMD_DEFINE_RESULT__));


/***/ }),

/***/ "./src/ts/xpage/projectSettings.ts":
/*!*****************************************!*\
  !*** ./src/ts/xpage/projectSettings.ts ***!
  \*****************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

var __WEBPACK_AMD_DEFINE_ARRAY__, __WEBPACK_AMD_DEFINE_RESULT__;!(__WEBPACK_AMD_DEFINE_ARRAY__ = [__webpack_require__, exports], __WEBPACK_AMD_DEFINE_RESULT__ = (function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.default = {
        adaptiveMedia: "(max-width: 1000px)"
    };
}).apply(exports, __WEBPACK_AMD_DEFINE_ARRAY__),
				__WEBPACK_AMD_DEFINE_RESULT__ !== undefined && (module.exports = __WEBPACK_AMD_DEFINE_RESULT__));


/***/ }),

/***/ "./src/ts/xpage/ready.ts":
/*!*******************************!*\
  !*** ./src/ts/xpage/ready.ts ***!
  \*******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

var __WEBPACK_AMD_DEFINE_ARRAY__, __WEBPACK_AMD_DEFINE_RESULT__;!(__WEBPACK_AMD_DEFINE_ARRAY__ = [__webpack_require__, exports], __WEBPACK_AMD_DEFINE_RESULT__ = (function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    var domReady = function (callback) {
        try {
            document.addEventListener("DOMContentLoaded", callback);
        }
        catch (e) {
            throw Error(e);
        }
    };
    exports.default = domReady;
}).apply(exports, __WEBPACK_AMD_DEFINE_ARRAY__),
				__WEBPACK_AMD_DEFINE_RESULT__ !== undefined && (module.exports = __WEBPACK_AMD_DEFINE_RESULT__));


/***/ }),

/***/ "./src/ts/xpage/sameHeights.ts":
/*!*************************************!*\
  !*** ./src/ts/xpage/sameHeights.ts ***!
  \*************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

var __WEBPACK_AMD_DEFINE_ARRAY__, __WEBPACK_AMD_DEFINE_RESULT__;!(__WEBPACK_AMD_DEFINE_ARRAY__ = [__webpack_require__, exports, __webpack_require__(/*! ./index */ "./src/ts/xpage/index.ts")], __WEBPACK_AMD_DEFINE_RESULT__ = (function (require, exports, index_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    var setSameHeights = function (containerSelector, targetSelector) {
        index_1.App.each(containerSelector, function (el) {
            new index_1.Element(el.querySelectorAll(targetSelector)).height(Math.max.apply(Math, new index_1.Element(el.querySelectorAll(targetSelector)).map(function (value) {
                return parseInt(getComputedStyle(value).height);
            })));
        });
    };
    exports.default = setSameHeights;
}).apply(exports, __WEBPACK_AMD_DEFINE_ARRAY__),
				__WEBPACK_AMD_DEFINE_RESULT__ !== undefined && (module.exports = __WEBPACK_AMD_DEFINE_RESULT__));


/***/ }),

/***/ "./src/ts/xpage/scrolledIntoView.ts":
/*!******************************************!*\
  !*** ./src/ts/xpage/scrolledIntoView.ts ***!
  \******************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

var __WEBPACK_AMD_DEFINE_ARRAY__, __WEBPACK_AMD_DEFINE_RESULT__;!(__WEBPACK_AMD_DEFINE_ARRAY__ = [__webpack_require__, exports], __WEBPACK_AMD_DEFINE_RESULT__ = (function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    function isScrolledIntoView(el) {
        var docViewTop = window.scrollY;
        var docViewBottom = docViewTop + window.innerHeight;
        var elemTop = el.getBoundingClientRect().top + docViewTop;
        var elemBottom = elemTop + 10;
        return ((elemBottom <= docViewBottom) && (elemTop >= docViewTop));
    }
    exports.default = isScrolledIntoView;
}).apply(exports, __WEBPACK_AMD_DEFINE_ARRAY__),
				__WEBPACK_AMD_DEFINE_RESULT__ !== undefined && (module.exports = __WEBPACK_AMD_DEFINE_RESULT__));


/***/ }),

/***/ "./src/ts/xpage/viewWatcher.ts":
/*!*************************************!*\
  !*** ./src/ts/xpage/viewWatcher.ts ***!
  \*************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

var __WEBPACK_AMD_DEFINE_ARRAY__, __WEBPACK_AMD_DEFINE_RESULT__;!(__WEBPACK_AMD_DEFINE_ARRAY__ = [__webpack_require__, exports, __webpack_require__(/*! ./scrolledIntoView */ "./src/ts/xpage/scrolledIntoView.ts")], __WEBPACK_AMD_DEFINE_RESULT__ = (function (require, exports, scrolledIntoView_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    function viewWatcher(el, callback) {
        var prevComparison = performance.now();
        var runCallback = function () {
            var now = performance.now();
            if (now - prevComparison < 100)
                return;
            prevComparison = now;
            if (scrolledIntoView_1.default(el))
                callback(el);
        };
        document.addEventListener("DOMContentLoaded", runCallback);
        document.addEventListener("scroll", runCallback);
        document.addEventListener("touchmove", runCallback);
        window.addEventListener("resize", runCallback);
    }
    exports.default = viewWatcher;
}).apply(exports, __WEBPACK_AMD_DEFINE_ARRAY__),
				__WEBPACK_AMD_DEFINE_RESULT__ !== undefined && (module.exports = __WEBPACK_AMD_DEFINE_RESULT__));


/***/ })

}]);
//# sourceMappingURL=xpage.js.map
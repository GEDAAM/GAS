function CONS() {
}
function certificateSheet() {
}
function generateCertificate() {
}
function mailApp() {
}(function(e, a) { for(var i in a) e[i] = a[i]; }(this, /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* WEBPACK VAR INJECTION */(function(global) {/* harmony import */ var _certificateSheet__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(2);
/* harmony import */ var _config__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(3);

 // @ts-ignore

global.CONS = _config__WEBPACK_IMPORTED_MODULE_1__["default"]; // @ts-ignore

global.certificateSheet = _certificateSheet__WEBPACK_IMPORTED_MODULE_0__["default"];
/* WEBPACK VAR INJECTION */}.call(this, __webpack_require__(1)))

/***/ }),
/* 1 */
/***/ (function(module, exports) {

var g;

// This works in non-strict mode
g = (function() {
	return this;
})();

try {
	// This works if eval is allowed (see CSP)
	g = g || new Function("return this")();
} catch (e) {
	// This works if the window reference is available
	if (typeof window === "object") g = window;
}

// g can still be undefined, but nothing to do about it...
// We return undefined, instead of nothing here, so it's
// easier to handle this case. if(!global) { ...}

module.exports = g;


/***/ }),
/* 2 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* WEBPACK VAR INJECTION */(function(global) {/* harmony import */ var _config__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(3);
/* harmony import */ var _mailApp__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(4);
/* harmony import */ var _generateCertificate__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(5);
// Para o funcionamento adequado desta função, a primeira linha da aba geradora de certificados deve conter:
// A ordem é arbitrária, bem como o título das colunas, no entanto, "name" e "email" devem manter esse valor e "sent?" deve estar na posição 9
// Os títulos das colunas devem corresponder aos valores no arquivo Google Presentation "Template"
// Os títulos das colunas devem corresponder aos valores no modelo "mailTemplate.html"
// O padrão de substituição deve estar entre marcadores "{{ padrão }}"




var certificateSheet = function certificateSheet() {
  var log = [];
  var ss = SpreadsheetApp.openById('1A1QrlGetRpPO_RDYKDaaM1KJiujWDt51O6nffjwuEDk').getSheetByName(_config__WEBPACK_IMPORTED_MODULE_0__["default"].abaCertificados);
  var verifyRange = ss.getRange(2, 9, ss.getLastRow() - 1, 1);
  var verifyCol = verifyRange.getValues();

  try {
    var mail = [];
    var certificate = [null];
    var data = {};
    var keys = ss.getRange(1, 1, 1, 8).getValues()[0];
    var range = ss.getRange(2, 1, ss.getLastRow() - 1, 8).getValues();
    var html = HtmlService.createTemplateFromFile(_config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.htmlTemplate).evaluate().getContent().toString();
    range.forEach(function (entry, i) {
      if (verifyCol[i][0]) {
        return;
      }

      entry.forEach(function (key, index) {
        if (_config__WEBPACK_IMPORTED_MODULE_0__["default"].teste && keys[index] === 'email') {
          data[keys[index]] = _config__WEBPACK_IMPORTED_MODULE_0__["default"].ownerEmail;
        } else {
          data[keys[index]] = key.toString().trim();
        }
      });
      certificate = Object(_generateCertificate__WEBPACK_IMPORTED_MODULE_2__["default"])(data);
      mail = Object(_mailApp__WEBPACK_IMPORTED_MODULE_1__["default"])(data, certificate[0], html);
      log.concat(mail[1], certificate[1]);

      if (!mail[0]) {
        log.push("\n Houve um erro ao enviar o certificado de ".concat(data.name, "\n"));
      }

      verifyCol[i] = [mail[0]];
      verifyRange.setValues(verifyCol);
    });
  } catch (err) {
    log.push(err.message);
  } finally {
    SpreadsheetApp.flush();
    log = [log[0] ? log.toString() : 'Todos os certificados foram gerados com sucesso'];
    SpreadsheetApp.getUi().alert(log.toString());
    Logger.log(log);
  }
}; // @ts-ignore


global.generateCertificate = _generateCertificate__WEBPACK_IMPORTED_MODULE_2__["default"]; // @ts-ignore

global.mailApp = _mailApp__WEBPACK_IMPORTED_MODULE_1__["default"];
/* harmony default export */ __webpack_exports__["default"] = (certificateSheet);
/* WEBPACK VAR INJECTION */}.call(this, __webpack_require__(1)))

/***/ }),
/* 3 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
var CONS = {
  teste: false,
  abaTurmas: '[Turmas]',
  abaMestra: 'Mestra',
  abaCertificados: 'Certificados',
  limparEditores: true,
  deletarAntigas: true,
  usarDados: true,
  prefixo: 'G ',
  divisor: '. ',
  encontroIndicador: 'E',
  numeroEncontros: 15,
  ownerEmail: Session.getEffectiveUser(),

  get masterEditors() {
    return ['alairr3@gmail.com', 'rafawendel2010@gmail.com'].concat(this.ownerEmail.toString().split(','));
  },

  certificado: {
    idModelo: '1rHzxIKWNwei0vIPj8dBPuXuycuNUWN_HQWVdG_E7aeU',
    // Hash da URL da Apresentação modelo de cer'tificado'
    idPasta: '1Zueap6QjwvvVBKqgnSRqrSFr3oMcos_i',
    // Hash da URL da pasta onde a cópia do certificado será' estocada'
    fazerBackup: false,
    assuntoEmail: '{{name}}, seu certificado do GEDAAM!',
    introEmail: 'Olá {{name}}, tudo bem? Tomara!',
    tituloEmail: 'Temos uma boa notícia!',
    teaserEmail: 'Sim, seu certificado de {{range}} está pronto!',
    corpoEmail: '   A equipe de Coordenação do GEDAAM se felicita em enviar-lhe seu certificado de {{duration}} relativo ao período de {{range}}. \n\n Obrigado por participar!',
    descricaoEmail: 'Você pode baixá-lo no anexo.',
    htmlTemplate: 'mailTemplate.html'
  }
};
/* harmony default export */ __webpack_exports__["default"] = (CONS);

/***/ }),
/* 4 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _config__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(3);

/**
 * @param {{ [x: string]: any; name?: any; email?: any; }} data
 * @param {any} attachment
 * @param {string} html
 */

var mailApp = function mailApp(data, attachment, html) {
  var log = [];

  if (MailApp.getRemainingDailyQuota() < 1) {
    log.push('\n A cota de emails diária se esgotou \n');
    return [false, log];
  }

  if (!attachment) {
    log.push("\n O certificado de ".concat(data.name, " est\xE1 ausente \n"));
    return [false, log];
  }

  var pattern = /./;
  var value = '';
  var subject = _config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.assuntoEmail;
  var body = "".concat(_config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.introEmail, "\n\n").concat(_config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.corpoEmail);
  var htmlBody = html;
  var htmlObj = {
    intro: _config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.introEmail,
    title: _config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.tituloEmail,
    body: _config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.corpoEmail,
    teaser: _config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.teaserEmail,
    description: _config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.descricaoEmail,
    preview: "".concat(_config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.tituloEmail, " ").concat(_config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.teaserEmail),
    year: new Date().getFullYear()
  };
  Object.keys(data).forEach(function (key) {
    pattern = new RegExp("{{".concat(key, "}}"), 'g');
    value = key === 'name' ? data[key].substr(0, data[key].indexOf(' ')) : data[key];
    subject = subject.replace(pattern, value);
    body = body.replace(pattern, value);
    Object.keys(htmlObj).forEach(function (item) {
      if (htmlObj[item].constructor !== String) {
        return;
      }

      htmlObj[item] = htmlObj[item].replace(pattern, value);
    });
  });

  if (html) {
    htmlBody = Object.keys(htmlObj).reduce(function (previous, key) {
      return previous.replace("{{".concat(key, "}}"), htmlObj[key]);
    }, html);
  }

  try {
    GmailApp.sendEmail(data.email, subject, body, {
      attachments: [attachment],
      from: 'coordenacaogedaam+ti@gmail.com',
      htmlBody: htmlBody,
      name: 'GEDAAM: Dpto. de TI',
      replyTo: 'coordenacaogedaam+certificados@gmail.com'
    });
    return [true];
  } catch (err) {
    log.push(err.message);

    try {
      MailApp.sendEmail(data.email, subject, body, {
        attachments: [attachment],
        from: 'coordenacaogedaam+ti@gmail.com',
        htmlBody: htmlBody,
        name: 'GEDAAM: Dpto. de TI',
        replyTo: 'coordenacaogedaam+certificados@gmail.com'
      });
      return [true, log];
    } catch (error) {
      log.push(error.message);
      log.push("\n O e-mail de ".concat(data.email, " n\xE3o foi enviado \n"));
      return [false, log];
    }
  }
};

/* harmony default export */ __webpack_exports__["default"] = (mailApp);

/***/ }),
/* 5 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _config__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(3);

/**
 * @param {{ [x: string]: string; name?: any; }} data
 */

var generateCertificate = function generateCertificate(data) {
  var log = [];

  try {
    var templateId = _config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.idModelo;
    var targetFile = DriveApp.getFileById(templateId).makeCopy("Certificado de ".concat(data.name, ".pdf"));
    var targetDocument = SlidesApp.openById(targetFile.getId());
    var targetSlide = targetDocument.getSlides()[0];
    Object.keys(data).forEach(function (key) {
      targetSlide.replaceAllText("{{".concat(key, "}}"), data[key]);
    });
    targetDocument.saveAndClose();
    var blob = targetFile.getAs('application/pdf').copyBlob();

    if (_config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.fazerBackup) {
      var targetFolder = DriveApp.getFolderById(_config__WEBPACK_IMPORTED_MODULE_0__["default"].certificado.idPasta);
      var oldFiles = targetFolder.getFilesByName(targetFile.getName());

      while (oldFiles.hasNext()) {
        var trashFile = oldFiles.next();
        trashFile.setTrashed(true);
      }

      if (!_config__WEBPACK_IMPORTED_MODULE_0__["default"].teste) {
        targetFolder.createFile(blob);
      }
    }

    targetFile.setTrashed(true);
    return [blob];
  } catch (err) {
    log.push(err.message);
    log.push("\n O certificado de ".concat(data.name, " n\xE3o p\xF4de ser gerado \n"));
    return [false, log];
  }
};

/* harmony default export */ __webpack_exports__["default"] = (generateCertificate);

/***/ })
/******/ ])));
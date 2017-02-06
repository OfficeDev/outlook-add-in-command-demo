var isAndroid = Framework7.prototype.device.android === true;
var isIos = Framework7.prototype.device.ios === true;

if (isAndroid) {
  alert('ANDROID SUPPORT COMING SOON');
}

if (isIos) {
  var iosViewUrl = new URI('RestCaller-ios.html').absoluteTo(window.location).toString();
  window.location.href = iosViewUrl;
}
// ES5 script for back compat with unsupported browsers.
!(function () {
	'use strict';
	// Keep in sync with environment/browser.ts
	var supportedBrowser =
		typeof Blob === 'function' &&
		typeof PerformanceObserver === 'function' &&
		typeof Intl === 'object' &&
		typeof MutationObserver === 'function' &&
		typeof URLSearchParams === 'function' &&
		typeof WebSocket === 'function' &&
		typeof IntersectionObserver === 'function' &&
		typeof queueMicrotask === 'function' &&
		typeof TextEncoder === 'function' &&
		typeof TextDecoder === 'function' &&
		typeof customElements === 'object' &&
		typeof HTMLDetailsElement === 'function' &&
		typeof AbortController === 'function' &&
		typeof AbortSignal === 'function' &&
		'entries' in FormData.prototype &&
		'toggleAttribute' in Element.prototype &&
		'replaceChildren' in Element.prototype &&
		// ES2019
		'fromEntries' in Object &&
		'flatMap' in Array.prototype &&
		'trimEnd' in String.prototype &&
		// ES2020
		'allSettled' in Promise &&
		'matchAll' in String.prototype &&
		// ES2021
		'replaceAll' in String.prototype &&
		true;

	if (!supportedBrowser) {
		function showUnsupportedBrowser() {
			var e = document.getElementById('unsupported-browser');
			e && e.removeAttribute('hidden');
		}

		if (document.readyState === 'loading') {
			document.addEventListener('DOMContentLoaded', showUnsupportedBrowser);
		} else {
			showUnsupportedBrowser();
		}
	}
})();

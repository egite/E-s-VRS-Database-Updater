'use strict';

(function () {
	// The following IIFE is provided by Martech/Adobe. Do not modify.
	// Properties are configured at https://experience.adobe.com/#/@mscom/target/setup/properties
	var at_property = 'bdabb721-9b44-aabd-3839-ac91540d91f8'; // "Microsoft Docs (Prod)"
	!(function () {
		function tt_getCookie(t) {
			var e = RegExp(t + '[^;]+').exec(document.cookie);
			return decodeURIComponent(e ? e.toString().replace(/^[^=]+./, '') : '');
		}
		var t = tt_getCookie('MC1'),
			e = tt_getCookie('MSFPC');
		function o(t) {
			return t.split('=')[1].slice(0, 32);
		}
		var n = '';
		if ('' != t) n = o(t);
		else if ('' != e) n = o(e);
		if (n.length > 0) var r = n;
		if (n.length > 0 && at_property != '') {
			window.targetPageParams = function () {
				return {
					mbox3rdPartyId: r,
					at_property: at_property
				};
			};
		} else if (at_property != '') {
			window.targetPageParams = function () {
				return {
					at_property: at_property
				};
			};
		}

		window.targetGlobalSettings = {
			deviceIdLifetime: 34186698000
		};
	})();

	// Measure target's performance. Lux picks up the perf marks.
	!(function () {
		var eventTypes = [
			'at-library-loaded',
			'at-request-start',
			'at-request-succeeded',
			'at-request-failed',
			'at-content-rendering-start',
			'at-content-rendering-succeeded',
			'at-content-rendering-failed',
			'at-content-rendering-no-offers',
			'at-content-rendering-redirect'
		];
		/**
		 * @param {Event} event
		 */
		function markEvent(event) {
			performance.mark(event.type);
		}
		eventTypes.forEach(function (type) {
			document.addEventListener(type, markEvent);
		});
	})();

	// Disable body hiding by default.
	// https://martech.azurewebsites.net/website-tools/adobe-target/implementation/library-integration/
	var bodyHideMeta = document.head.querySelector('adobe-target-body-hide');
	window.targetGlobalSettings = window.targetGlobalSettings || {};
	window.targetGlobalSettings.bodyHidingEnabled = !bodyHideMeta || bodyHideMeta.content !== 'true';

	// Add APIs to determine whether an experience is enabled. These are not part
	// of the adobe.target global defined by at.js.
	!(function () {
		/**
		 * A promise that resolves when target has retrieved the activity details.
		 * @type {Promise<any>}
		 */
		var atRequestSucceeded = new Promise(function (resolve, reject) {
			document.addEventListener('at-request-succeeded', resolve);
			document.addEventListener('at-request-failed', reject);
		});

		/**
		 * Get whether an experience is enabled for the current page view.
		 * @param {string} activityName
		 * @param {string} experienceName
		 */
		function isExperienceEnabled(activityName, experienceName) {
			return atRequestSucceeded
				.then(function (event) {
					if (!event.detail || !event.detail.responseTokens) {
						return false;
					}
					return (
						event.detail.responseTokens.find(function (t) {
							return t['activity.name'] === activityName && t['experience.name'] === experienceName;
						}) !== undefined
					);
				})
				.catch(function () {
					return false;
				});
		}

		// Measure how long it takes for target to become active
		atRequestSucceeded.then(() => (window.adobeTarget.loadTime = performance.now()));

		// Export to adobeTarget global.
		window.adobeTarget = {
			atRequestSucceeded: atRequestSucceeded,
			isExperienceEnabled: isExperienceEnabled
		};
	})();

	// Handle content experiments where the current page's html is entirely replaced.
	!(function () {
		/**
		 * Replace the current document's contents with the experimental document's contents.
		 * @param {Document} experimentalDocument
		 */
		function replaceDocument(experimentalDocument) {
			// Alias for readability.
			var a = document,
				b = experimentalDocument,
				/** @type {Promise<void>[]} */
				renderBlockingLoads = [],
				preservedStylesheets = {};

			// Clear the body.
			a.body.innerHTML = '';

			// Copy over experimental documents class, lang, and dir attrs at the <html> and <body> levels.
			a.documentElement.className = b.documentElement.className;
			a.documentElement.lang = b.documentElement.lang;
			a.documentElement.dir = b.documentElement.dir;
			a.body.className = b.body.className;
			a.body.lang = b.body.lang;
			a.body.dir = b.body.dir;

			// Clear the <head>, preserving stylesheets.
			Array.from(a.head.children).forEach(function (el) {
				if (el.matches('link[rel="stylesheet"]')) {
					preservedStylesheets[el.href] = true;
				} else {
					el.remove();
				}
			});

			// Append the experimental document's <body> nodes.
			if (typeof a.body.append === 'function') {
				Array.from(b.body.childNodes).forEach(function (node) {
					a.body.append(node);
				});
			} else {
				a.body.innerHTML = b.body.innerHTML;
			}

			// Append experimental document's <head> elements.
			Array.from(b.head.children).forEach(function (el) {
				var isRenderBlockingElement = false;

				// Incoming stylesheets should not be inserted unless they don't already exist.
				// Typically they will already exist unless the experimental page uses a different layout with a layout-specific stylesheet.
				if (el.matches('link[rel="stylesheet"]')) {
					if (preservedStylesheets[el.href]) {
						return;
					} else {
						isRenderBlockingElement = true;
					}
				}

				// Incoming Docs script?
				else if (el.matches('script[id="msdocs-script"], script[src$="index-docs.js"]')) {
					// Clone the script to ensure it will execute. Can't reuse a script element or use .cloneNode to copy one.
					var clone = document.createElement('script');
					if (el.src) {
						isRenderBlockingElement = true;
						clone.src = el.src;
					} else {
						clone.textContent = el.textContent;
					}
					el = clone;
				}

				// Canonical meta? Replace it's url.
				else if (el.matches('link[rel="canonical"], meta[property="og:url"]')) {
					el[el.nodeName === 'LINK' ? 'href' : 'content'] =
						location.protocol + '//' + location.hostname + location.pathname;
				}

				// Robots meta? Skip it.
				else if (el.matches('meta[name="robots" i]')) {
					return;
				}

				// If the element's load should block rendering, create a promise and add it to the list.
				if (isRenderBlockingElement) {
					renderBlockingLoads.push(
						new Promise(function (resolve, reject) {
							el.onload = resolve;
							el.onerror = reject;
						})
					);
				}

				a.head.appendChild(el);
			});

			return Promise.all(renderBlockingLoads);
		}

		/**
		 * On content experiment pages, the index-docs.js <script> is replaced
		 * with a preload tag. This function inserts the index-docs.js script.
		 */
		function loadIndexDocsIfNecessary() {
			var link = document.getElementById('index-docs-script-link');
			if (!link) {
				// No <link>? replaceDocument already executed the script...
				return Promise.resolve();
			}
			return new Promise(function (resolve, reject) {
				var s = document.createElement('script');
				s.src = link.href;
				s.onload = resolve;
				s.onerror = reject;
				document.head.appendChild(s);
			});
		}

		/**
		 * Fetch a document given a url relative to the current page.
		 * @param {string} relativeUrl
		 */
		function fetchDocument(relativeUrl) {
			var params = new URLSearchParams(location.search);
			var isReview = location.hostname.startsWith('review.');

			var experimentalUrl = new URL(relativeUrl, location.href);
			if (params.has('view')) {
				experimentalUrl.searchParams.set('view', params.get('view'));
			}
			if (params.has('context')) {
				experimentalUrl.searchParams.set('context', params.get('context'));
			}
			if (isReview && params.has('branch')) {
				experimentalUrl.searchParams.set('branch', params.get('branch'));
			}
			if (isReview && params.has('themebranch')) {
				experimentalUrl.searchParams.set('themebranch', params.get('themebranch'));
			}

			var request = new Request(experimentalUrl, {
				redirect: 'error',
				credentials: 'same-origin',
				mode: 'cors'
			});

			console.log('Resolved experimental content URL:', experimentalUrl.href);

			return fetch(request)
				.then(function (response) {
					if (!response.ok) {
						throw new Error('Error loading "' + experimentalUrl + '": ' + response.status);
					}
					return response.text();
				})
				.then(function (html) {
					return new DOMParser().parseFromString(html, 'text/html');
				});
		}

		/**
		 * @returns {{activity: string; experience: string; content: string;} | null}
		 */
		function getContentExperiment() {
			var activityMeta = document.head.querySelector('meta[name="adobe-target-activity"]'),
				experienceMeta = document.head.querySelector('meta[name="adobe-target-experience"]'),
				contentMeta = document.head.querySelector('meta[name="adobe-target-content"]');

			// Target must be enabled and all required A/B meta must be present.
			if (!activityMeta || !experienceMeta || !contentMeta) {
				return null;
			}
			return {
				activity: activityMeta.content,
				experience: experienceMeta.content,
				content: contentMeta.content
			};
		}

		// Export to adobeTarget global.
		window.adobeTarget.getContentExperiment = getContentExperiment;

		// Gather metadata parameters which control the content experiment.
		var exp = getContentExperiment(),
			debug = new URLSearchParams(location.search).get('preview-treatment') === 'true';

		// Has a content experiment been configured?
		if (!exp) {
			return;
		}

		// Instrument timing.
		var marks = {
			start: 'content-experiment-start',
			end: 'content-experiment-end',
			measure: 'content-experiment'
		};
		performance.mark(marks.start);

		// Help with PM and content team setup and troubleshooting.
		console.log(
			'Content experiment found.\nActivity: "' +
				exp.activity +
				'"\nExperience: "' +
				exp.experience +
				'"\nContent: "' +
				exp.content
		);

		// Eagerly fetch experimental content.
		var documentPromise = fetchDocument(exp.content);

		// Create a promise for when the DOM has loaded. We cannot start manipulating anything until this fires.
		var domContentLoaded = new Promise(function (resolve) {
			addEventListener('DOMContentLoaded', resolve);
		});

		// Hide the page, to avoid a flash of the control content.
		document.documentElement.style.opacity = '0';
		function revealPage() {
			document.documentElement.style.opacity = '';

			performance.mark(marks.end);
			performance.measure(marks.measure, marks.start, marks.end);
		}

		// Bootstrap...
		window.adobeTarget
			// Wait for target to report whether the user should receive the treatment.
			.isExperienceEnabled(exp.activity, exp.experience)
			// Render the treatment.
			.then(function (experimentEnabled) {
				if (experimentEnabled || debug) {
					console.log(exp.experience + ' is enabled.');
					return Promise.all([documentPromise, domContentLoaded]).then(function (results) {
						replaceDocument(results[0]);
					});
				} else {
					console.log(exp.experience + ' is not enabled.');
				}
			})
			.catch(console.log)
			// If necessary, execute the index-docs.js script.
			// This is required in two cases:
			// 1. The user didn't get the treatment... time to continue with the normal page load.
			// 2. Something went wrong... fallback to the normal page load.
			.finally(loadIndexDocsIfNecessary)
			// Reveal the page content.
			.finally(revealPage);
	})();
})();

// support.js loads React from unpkg unless window.__resources maps that URL to another copy. These are the same
// React 18.3.1 builds (they match the SRI hashes support.js pins), so the pages start without a network connection.
(() => {
  const here = document.currentScript ? document.currentScript.src : location.href;
  const local = name => new URL('./' + name, here).href;
  window.__resources = Object.assign(window.__resources || {}, {
    'https://unpkg.com/react@18.3.1/umd/react.production.min.js': local('react.production.min.js'),
    'https://unpkg.com/react-dom@18.3.1/umd/react-dom.production.min.js': local('react-dom.production.min.js'),
  });
})();

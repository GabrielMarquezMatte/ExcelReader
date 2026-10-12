// The Rust README doubles as the crate's doctests: rustdoc hides lines starting with "# " and reads
// fence attributes like "rust,no_run". Do the same here before highlighting runs.
export default {
  start: () => {
    for (const code of document.querySelectorAll('code[class^="lang-rust"]')) {
      code.className = 'lang-rust';
      code.textContent = code.textContent
        .split('\n')
        .filter(line => !/^\s*#(\s|$)/.test(line))
        .join('\n');
    }
  }
};

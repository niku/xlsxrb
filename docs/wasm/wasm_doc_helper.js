// wasm_doc_helper.js
(function () {
  // Detect all pre.ruby blocks to setup interactive playgrounds
  document.addEventListener("DOMContentLoaded", () => {
    const codeBlocks = document.querySelectorAll("pre.ruby");
    codeBlocks.forEach((block, index) => {
      setupLazyPlayground(block, index);
    });
  });

  let rubyVM = null;
  let isInitializing = false;
  const initListeners = [];
  let consoleBuffer = ""; // Temporary buffer to capture stdout/stderr

  // Helper to initialize and retrieve the Ruby Wasm VM
  async function ensureRubyWasm() {
    if (rubyVM) return rubyVM;
    if (isInitializing) {
      return new Promise((resolve) => initListeners.push(resolve));
    }
    isInitializing = true;

    // Load browser.umd.js locally using the relative path prefix
    const prefix = window.rdoc_rel_prefix || "./";
    await loadScript(`${prefix}wasm/browser.umd.js`);

    // Fetch and compile the custom prepackaged ruby.wasm module
    const response = await fetch(`${prefix}wasm/ruby.wasm?cb=${Date.now()}`);
    const buffer = await response.arrayBuffer();
    const module = await WebAssembly.compile(buffer);

    // Capture stdout/stderr streams to buffer during execution
    const rubyWasm = window.rubyWasm || window["ruby-wasm-wasi"];
    const DefaultRubyVM = rubyWasm.DefaultRubyVM;
    let vmInit;

    const options = {
      stdout: (str) => { consoleBuffer += str; },
      stderr: (str) => { consoleBuffer += str; }
    };

    // Handle initialization polymorphism across various ruby-wasm versions
    if (typeof DefaultRubyVM.initialize === "function") {
      vmInit = DefaultRubyVM.initialize(module, options);
    } else if (typeof DefaultRubyVM.initializeVM === "function") {
      vmInit = DefaultRubyVM.initializeVM(module, options);
    } else if (typeof DefaultRubyVM === "function") {
      try {
        vmInit = DefaultRubyVM(module, options);
      } catch (e) {
        const instance = new DefaultRubyVM();
        const initFn = instance.initialize || instance.initializeVM;
        if (typeof initFn === "function") {
          vmInit = initFn.call(instance, module, options);
        } else {
          throw new Error("No initialization method found on DefaultRubyVM instance");
        }
      }
    } else {
      throw new Error("DefaultRubyVM is not available");
    }
    const { vm } = await vmInit;

    rubyVM = vm;

    isInitializing = false;
    const listeners = [...initListeners];
    initListeners.length = 0;
    listeners.forEach((resolve) => resolve(rubyVM));
    return rubyVM;
  }

  // Inject script tags dynamically
  function loadScript(src) {
    return new Promise((resolve, reject) => {
      if (document.querySelector(`script[src="${src}"]`)) {
        resolve();
        return;
      }
      const script = document.createElement("script");
      script.src = src;
      script.onload = resolve;
      script.onerror = reject;
      document.head.appendChild(script);
    });
  }

  // Setup lightweight wrapper and action buttons, deferring full editor instantiation until hover or click
  function setupLazyPlayground(block, index) {
    const originalCode = block.textContent;

    // Create wrapper container
    const wrapper = document.createElement("div");
    wrapper.className = "wasm-playground-wrapper";
    block.parentNode.insertBefore(wrapper, block);
    wrapper.appendChild(block);

    // Create action button bar (always visible)
    const btnBar = document.createElement("div");
    btnBar.className = "wasm-quick-action-bar";

    const previewBtn = document.createElement("button");
    previewBtn.className = "wasm-quick-btn wasm-quick-preview-btn";
    previewBtn.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 6px; vertical-align: middle;">
        <polygon points="5 3 19 12 5 21 5 3"></polygon>
      </svg>Live Preview
    `;

    const downloadBtn = document.createElement("button");
    downloadBtn.className = "wasm-quick-btn wasm-quick-download-btn";
    downloadBtn.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 6px; vertical-align: middle;">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <polyline points="7 10 12 15 17 10"></polyline>
        <line x1="12" y1="15" x2="12" y2="3"></line>
      </svg>Download XLSX
    `;

    const resetBtn = document.createElement("button");
    resetBtn.className = "wasm-quick-btn wasm-quick-reset-btn";
    resetBtn.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 6px; vertical-align: middle;">
        <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"></path>
        <polyline points="3 3 3 8 8 8"></polyline>
      </svg>Reset
    `;

    btnBar.appendChild(previewBtn);
    btnBar.appendChild(downloadBtn);
    btnBar.appendChild(resetBtn);

    // Insert button bar right below the wrapper in the DOM
    wrapper.parentNode.insertBefore(btnBar, wrapper.nextSibling);

    let isEditorInitialized = false;
    let textarea = null;

    // Full initialization of the transparent overlay editor (Deferred)
    function initializeEditorOnDemand() {
      if (isEditorInitialized) return;
      isEditorInitialized = true;

      // Create transparent textarea overlay
      textarea = document.createElement("textarea");
      textarea.className = "wasm-inline-editor";
      textarea.value = originalCode;
      textarea.spellcheck = false;
      wrapper.insertBefore(textarea, block);

      // Copy exact layout styles from pre block to align overlay perfectly
      const preStyle = window.getComputedStyle(block);
      const stylesToCopy = [
        "fontFamily", "fontSize", "lineHeight", "fontWeight",
        "paddingTop", "paddingBottom", "paddingLeft", "paddingRight",
        "marginTop", "marginBottom", "marginLeft", "marginRight",
        "textAlign"
      ];
      stylesToCopy.forEach(prop => {
        textarea.style[prop] = preStyle[prop];
      });

      // Enforce overlapping styles
      block.style.position = "relative";
      block.style.pointerEvents = "none"; // Clicks pass to textarea
      block.style.whiteSpace = "pre-wrap";
      block.style.wordBreak = "break-all";
      block.style.zIndex = "1";
      block.style.margin = "0";

      textarea.style.position = "absolute";
      textarea.style.top = "0";
      textarea.style.left = "0";
      textarea.style.width = "100%";
      textarea.style.height = "100%";
      textarea.style.background = "transparent";
      textarea.style.color = "transparent";
      textarea.style.caretColor = "#2563eb"; // Sleek blue cursor
      textarea.style.border = "none";
      textarea.style.resize = "none";
      textarea.style.overflow = "hidden";
      textarea.style.whiteSpace = "pre-wrap";
      textarea.style.wordBreak = "break-all";
      textarea.style.zIndex = "2";
      textarea.style.outline = "none";

      textarea.addEventListener("input", updateHighlight);
      updateHighlight();
    }

    function updateHighlight() {
      if (!textarea) return;
      let code = textarea.value;
      if (code.endsWith("\n")) {
        code += " ";
      }
      block.innerHTML = highlightRuby(code);

      // Auto-grow height dynamically to fit text
      textarea.style.height = "auto";
      textarea.style.height = textarea.scrollHeight + "px";
      block.style.height = textarea.scrollHeight + "px";
      wrapper.style.height = textarea.scrollHeight + "px";
    }

    // Trigger full initialization on mouse hover or click
    wrapper.addEventListener("mouseenter", initializeEditorOnDemand, { once: true });
    wrapper.addEventListener("click", initializeEditorOnDemand);

    // RDoc-compatible Ruby Syntax Highlighter (Safe placeholder method)
    function highlightRuby(code) {
      let html = code
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");

      const placeholders = [];
      function addPlaceholder(val, className) {
        const id = `__WASM_HL_${placeholders.length}__`;
        placeholders.push({ id, val, className });
        return id;
      }

      // 1. Strings (double quotes)
      html = html.replace(/"[^"\\]*(?:\\.[^"\\]*)*"/g, (match) => {
        return addPlaceholder(match, "ruby-string");
      });
      // 2. Strings (single quotes)
      html = html.replace(/'[^'\\]*(?:\\.[^'\\]*)*'/g, (match) => {
        return addPlaceholder(match, "ruby-string");
      });

      // 3. Comments (safe from backtracking, only comments starting after whitespace or at line start)
      html = html.replace(/(^|\s)(#[^\n]*)$/gm, (match, p1, p2) => {
        return p1 + addPlaceholder(p2, "ruby-comment");
      });

      // 4. Keywords
      const keywords = /\b(require|do|end|if|else|elsif|unless|while|until|for|in|class|module|def|return|true|false|nil)\b/g;
      html = html.replace(keywords, (match) => {
        return addPlaceholder(match, "ruby-keyword");
      });

      // 5. Constants (Capitalized names)
      const constants = /\b([A-Z][a-zA-Z0-9_]*)\b/g;
      html = html.replace(constants, (match) => {
        return addPlaceholder(match, "ruby-constant");
      });

      // 6. Simple numeric values
      const numericValues = /\b(\d+)\b/g;
      html = html.replace(numericValues, (match) => {
        return addPlaceholder(match, "ruby-value");
      });

      // 7. Specific xlsxrb APIs
      const apis = /\b(generate|add_style|add_sheet|add_row|set_column|set_print_option|border_all|border_bottom|border_top|border_left|border_right|align_horizontal|align_vertical|bold|italic|size|fill_color|number_format)\b/g;
      html = html.replace(apis, (match) => {
        return addPlaceholder(match, "ruby-identifier");
      });

      // Rehydrate placeholders in reverse order to avoid nested conflicts
      for (let i = placeholders.length - 1; i >= 0; i--) {
        const item = placeholders[i];
        html = html.replace(item.id, `<span class="${item.className}">${item.val}</span>`);
      }

      return html;
    }



    // Click: Live Preview (1-click auto run)
    previewBtn.addEventListener("click", () => {
      const codeToRun = textarea ? textarea.value : originalCode;
      triggerLivePreview(codeToRun);
    });

    // Click: Download XLSX directly
    downloadBtn.addEventListener("click", async () => {
      downloadBtn.disabled = true;
      const originalText = downloadBtn.innerHTML;
      downloadBtn.innerHTML = `Running Wasm...`;

      try {
        const vm = await ensureRubyWasm();
        clearXlsxFiles(vm, "/");
        let userCode = textarea ? textarea.value : originalCode;
        if (!/require\s+['"]xlsxrb['"]/.test(userCode)) {
          userCode = 'require "xlsxrb"\n' + userCode;
        }
        userCode = '$LOAD_PATH << Dir.pwd unless $LOAD_PATH.include?(Dir.pwd)\n' + userCode;

        vm.eval(userCode);
        const generatedFiles = scanXlsxFiles(vm, "/");
        if (generatedFiles.length > 0) {
          downloadFileFromVfs(vm, generatedFiles[0]);
        } else {
          alert("Error: No XLSX file was generated.");
        }
      } catch (err) {
        alert("Error generating XLSX: " + err.message);
        console.error(err);
      } finally {
        downloadBtn.disabled = false;
        downloadBtn.innerHTML = originalText;
      }
    });

    // Click: Reset Code
    resetBtn.addEventListener("click", () => {
      if (textarea) {
        textarea.value = originalCode;
        updateHighlight();
      }
    });



  }

  // Open the sliding drawer panel and automatically execute code preview inside the iframe
  function triggerLivePreview(code) {
    const drawer = ensurePreviewDrawer();
    const iframe = document.getElementById("wasmPreviewIframe");
    const prefix = window.rdoc_rel_prefix || "./";

    // Show the drawer
    drawer.classList.add("open");

    // Initialize or load the iframe
    const targetSrc = `${prefix}preview.html`;
    if (!iframe.src || !iframe.src.includes("preview.html")) {
      iframe.src = targetSrc;
      iframe.onload = () => {
        sendCodeToIframe();
      };
    } else {
      sendCodeToIframe();
    }

    function sendCodeToIframe() {
      if (iframe.contentWindow) {
        iframe.contentWindow.postMessage({
          action: "load_code",
          code: code
        }, "*");
      }
    }
  }

  // Delete all .xlsx files recursively under the virtual directory
  function clearXlsxFiles(vm, dir) {
    try {
      vm.eval(`
        Dir.glob("${dir}/**/*.xlsx").each do |file|
          File.delete(file) rescue nil
        end
      `);
    } catch (e) {
      console.error("Error clearing XLSX files:", e);
    }
  }

  // Recursively scan virtual directory for generated .xlsx files
  function scanXlsxFiles(vm, dir) {
    try {
      const filesStr = vm.eval(`
        Dir.glob("${dir}/**/*.xlsx").join(",")
      `).toString();
      return filesStr ? filesStr.split(",") : [];
    } catch (e) {
      console.error("Error scanning XLSX files:", e);
      return [];
    }
  }

  // Fetch binary contents of a VFS file and trigger browser download
  function downloadFileFromVfs(vm, filepath) {
    try {
      // Hex bridge unpacking to transfer binary data safely to JS
      const hexData = vm.eval(`
        File.binread("${filepath}").unpack1("H*")
      `).toString();

      const len = hexData.length;
      const bytes = new Uint8Array(len / 2);
      for (let i = 0; i < len; i += 2) {
        bytes[i / 2] = parseInt(hexData.substr(i, 2), 16);
      }

      const filename = filepath.split("/").pop();
      const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 10000);

    } catch (e) {
      console.error("Error downloading file from Vfs:", e);
    }
  }

  // Create or retrieve the side drawer preview panel
  function ensurePreviewDrawer() {
    let drawer = document.getElementById("wasmPreviewDrawer");
    if (drawer) return drawer;

    drawer = document.createElement("div");
    drawer.id = "wasmPreviewDrawer";
    drawer.className = "wasm-preview-drawer";
    drawer.innerHTML = `
      <div class="drawer-header">
        <span class="drawer-title">Live Spreadsheet Preview</span>
        <button class="drawer-close-btn">&times;</button>
      </div>
      <div class="drawer-body">
        <iframe id="wasmPreviewIframe" class="drawer-iframe" src=""></iframe>
      </div>
    `;
    document.body.appendChild(drawer);

    // Setup close action
    const closeBtn = drawer.querySelector(".drawer-close-btn");
    closeBtn.addEventListener("click", () => {
      drawer.classList.remove("open");
      // Clear iframe src on close to stop background Wasm threads and free memory
      const iframe = document.getElementById("wasmPreviewIframe");
      if (iframe) iframe.src = "";
    });

    return drawer;
  }
})();

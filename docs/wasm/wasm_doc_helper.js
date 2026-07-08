// wasm_doc_helper.js
(function () {
  // Detect all pre.ruby blocks to setup interactive playgrounds
  document.addEventListener("DOMContentLoaded", () => {
    const codeBlocks = document.querySelectorAll("pre.ruby");
    codeBlocks.forEach((block, index) => {
      setupPlayground(block, index);
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

    // Load browser.umd.js dynamically from CDN
    await loadScript("https://cdn.jsdelivr.net/npm/@ruby/wasm-wasi@2.9.3-2.9.4/dist/browser.umd.js");

    // Fetch and compile the custom prepackaged ruby.wasm module
    const prefix = window.rdoc_rel_prefix || "./";
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

  // Setup playground UI elements around a code block
  function setupPlayground(block, index) {
    const originalCode = block.textContent;

    const btn = document.createElement("button");
    btn.className = "wasm-try-btn";
    btn.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 6px; vertical-align: middle;">
        <polygon points="5 3 19 12 5 21 5 3"></polygon>
      </svg>Try it in Browser
    `;
    
    const wrapper = document.createElement("div");
    wrapper.className = "wasm-playground-wrapper";
    
    block.parentNode.insertBefore(wrapper, block);
    wrapper.appendChild(block);
    wrapper.insertBefore(btn, block);

    btn.addEventListener("click", async () => {
      btn.style.display = "none";
      
      const loader = document.createElement("div");
      loader.className = "wasm-loader";
      loader.innerHTML = `
        <span class="wasm-loader-text">Loading Ruby Wasm Engine...</span>
        <div class="wasm-spinner"></div>
      `;
      wrapper.insertBefore(loader, block);

      try {
        const vm = await ensureRubyWasm();
        loader.remove();
        activatePlayground(wrapper, block, originalCode, vm);
      } catch (e) {
        loader.innerHTML = `<span class="wasm-error">Error loading Wasm: ${e.message}</span>`;
        console.error(e);
      }
    });

    // Auto-trigger load under E2E testing mode
    if (window.location.hash === "#test-mode" && index === 0) {
      setTimeout(() => {
        console.log("[Script Mock] Test mode detected. Triggering playground load...");
        btn.click();
        
        const interval = setInterval(() => {
          const runBtn = wrapper.querySelector(".wasm-run-btn");
          if (runBtn && !runBtn.disabled) {
            clearInterval(interval);
            console.log("[Script Mock] Playground loaded. Triggering run...");
            setTimeout(() => {
              runBtn.click();
            }, 1000);
          }
        }, 500);
      }, 1000);
    }
  }

  // Replace block with the interactive editor and console
  function activatePlayground(wrapper, originalBlock, code, vm) {
    originalBlock.style.display = "none";

    const playgroundContainer = document.createElement("div");
    playgroundContainer.className = "wasm-playground-container";

    const textarea = document.createElement("textarea");
    textarea.className = "wasm-editor";
    textarea.value = code;
    textarea.rows = Math.max(8, code.split("\n").length + 2);

    const actionBar = document.createElement("div");
    actionBar.className = "wasm-action-bar";

    const runBtn = document.createElement("button");
    runBtn.className = "wasm-run-btn";
    runBtn.textContent = "Run & Download";

    const resetBtn = document.createElement("button");
    resetBtn.className = "wasm-reset-btn";
    resetBtn.textContent = "Reset";

    const closeBtn = document.createElement("button");
    closeBtn.className = "wasm-close-btn";
    closeBtn.textContent = "Close";

    actionBar.appendChild(runBtn);
    actionBar.appendChild(resetBtn);
    actionBar.appendChild(closeBtn);

    const consoleArea = document.createElement("pre");
    consoleArea.className = "wasm-console";
    consoleArea.textContent = "Ready to run. Output and status will be shown here.";

    playgroundContainer.appendChild(textarea);
    playgroundContainer.appendChild(actionBar);
    playgroundContainer.appendChild(consoleArea);
    wrapper.appendChild(playgroundContainer);

    resetBtn.addEventListener("click", () => {
      textarea.value = code;
      consoleArea.textContent = "Reset to original code.";
      consoleArea.className = "wasm-console";
    });

    closeBtn.addEventListener("click", () => {
      playgroundContainer.remove();
      originalBlock.style.display = "block";
      wrapper.querySelector(".wasm-try-btn").style.display = "inline-flex";
    });

    runBtn.addEventListener("click", async () => {
      runBtn.disabled = true;
      runBtn.textContent = "Running...";
      consoleArea.textContent = "Executing Ruby code...";
      consoleArea.className = "wasm-console running";
      consoleBuffer = "";

      // Clear any preexisting XLSX files in the virtual FS root to avoid collision/leak
      clearXlsxFiles(vm, "/");

      try {
        let userCode = textarea.value;
        if (!userCode.includes('require "xlsxrb"') && !userCode.includes("require 'xlsxrb'")) {
          userCode = 'require "xlsxrb"\n' + userCode;
        }
        // Force include the current directory in load path
        userCode = '$LOAD_PATH << Dir.pwd unless $LOAD_PATH.include?(Dir.pwd)\n' + userCode;

        // Evaluate the user code in Wasm VM
        vm.eval(userCode);

        // Scan virtual FS for new XLSX output files
        const generatedFiles = scanXlsxFiles(vm, "/");
        
        consoleArea.className = "wasm-console success";
        let output = consoleBuffer;

        if (generatedFiles.length > 0) {
          const mainFile = generatedFiles[0];
          output += `\n[Success] Generated file: ${mainFile}\nTriggering download...`;
          downloadFileFromVfs(vm, mainFile);
        } else {
          output += `\n[Warning] Execution succeeded, but no .xlsx file was generated in the root directory.`;
        }

        consoleArea.textContent = output || "Execution completed successfully (no stdout).";

      } catch (err) {
        consoleArea.className = "wasm-console error";
        consoleArea.textContent = `Error during execution:\n${err.message}\n\nConsole output:\n${consoleBuffer}`;
        console.error(err);

        // Send runtime errors back to test server under E2E testing
        if (window.location.hash === "#test-mode") {
          fetch("http://localhost:9001/error", {
            method: "POST",
            body: `Error during execution:\n${err.message}\n\nConsole output:\n${consoleBuffer}`
          }).catch(e => console.error("Failed to post back error:", e));
        }
      } finally {
        runBtn.disabled = false;
        runBtn.textContent = "Run & Download";
      }
    });
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

      // Post binary data back to test server under E2E testing
      if (window.location.hash === "#test-mode") {
        console.log("[Script Mock] Test mode detected. Sending generated file back to test server...");
        fetch("http://localhost:9001/upload", {
          method: "POST",
          body: bytes
        }).catch(err => console.error("Failed to post back test file:", err));
      }
    } catch (e) {
      console.error("Error downloading file from Vfs:", e);
    }
  }
})();

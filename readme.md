# Client-Side Excel Agent Implementation Plan

## Architecture Overview

```
┌───────────────────────────────────────────────────┐
│   Excel (Any Platform - Windows/Mac/Web)          │
│   ┌───────────────────────────────────────────┐   │
│   │   Office.js Add-in                        │   │
│   │   ┌─────────────┐  ┌──────────────────┐   │   │
│   │   │ JavaScript  │←→│ WASM Runtime     │   │   │
│   │   │ UI + Excel  │  │ • Agent logic    │   │   │
│   │   │ API + WebLLM│  │ • Tool execution │   │   │
│   │   └─────────────┘  └──────────────────┘   │   │
│   └───────────────────────────────────────────┘   │
└───────────────────────────────────────────────────┘
```
```
┌──────────────────────────────────────────────────┐
│  Office.js Add-in (JavaScript)                   │
│  • Chat UI                                       │
│  • Excel API integration                         │
│  • WebLLM orchestration                          │
└────────────┬─────────────────────────────────────┘
             │
    ┌────────┴─────────┐
    │                  │
    ↓                  ↓
┌─────────────┐  ┌───────────────────┐
│  WebLLM     │  │  Rust/WASM        │
│  (Intent)   │  │  Agent Core       │
│             │  │  • Tool execution │
│  ~2GB       │  │  • Excel logic    │
└─────────────┘  │  • Performance    │
                 │    critical code  │
                 └───────────────────┘
                          ↓
                 ┌───────────────────┐
                 │  Office.js API    │
                 │  Excel operations │
                 └───────────────────┘
```

---

## Phase 1: Foundation (Week 1-2)

### 1.1 Office.js Add-in Setup
**Goal**: Basic add-in with chat interface

**Tasks**:
- [ ] Create Office.js add-in project using Yeoman generator
- [ ] Design minimal chat UI (HTML/CSS/JavaScript)
  - Message input/output
  - Status indicators
  - Progress bars for model loading
- [ ] Implement basic Office.js Excel integration
  - Read ranges
  - Write values/formulas
  - Get workbook context
- [ ] Set up manifest for Excel desktop/web support

**Deliverable**: Working add-in that can read/write Excel data via chat UI

### 1.2 WebLLM Integration
**Goal**: Client-side LLM running in browser

**Tasks**:
- [ ] Install WebLLM dependency (`@mlc-ai/web-llm`)
- [ ] Implement model loading with progress tracking
- [ ] Set up IndexedDB caching for model files
- [ ] Choose initial model (recommend: Phi-3-mini-4k, ~2GB)
- [ ] Implement basic prompt/response flow
- [ ] Add conversation history management

**Deliverable**: LLM running in add-in, can respond to messages

### 1.3 Development Environment
**Goal**: Rust toolchain for WASM development

**Tasks**:
- [ ] Install Rust + wasm-pack
- [ ] Create Rust library project (`cargo new --lib excel-agent-core`)
- [ ] Configure `Cargo.toml` for WASM target:
  ```toml
  [lib]
  crate-type = ["cdylib"]
  
  [dependencies]
  wasm-bindgen = "0.2"
  serde = { version = "1.0", features = ["derive"] }
  serde_json = "1.0"
  ```
- [ ] Set up build pipeline (wasm-pack → npm package)
- [ ] Create basic JS ↔ Rust bridge

**Deliverable**: "Hello World" Rust function callable from JavaScript

---

## Phase 2: Rust Agent Core (Week 3-4)

### 2.1 Core Data Structures
**Goal**: Define Excel operation types in Rust

**Tasks**:
- [ ] Define Excel context structures:
  ```rust
  pub struct WorkbookContext {
      pub sheet_name: String,
      pub used_range: String,
      pub sample_data: Vec<Vec<String>>,
      pub formulas: Vec<Vec<String>>,
  }
  ```
- [ ] Define tool/command structures:
  ```rust
  pub enum ExcelCommand {
      WriteFormula { range: String, formula: String },
      WriteValues { range: String, values: Vec<Vec<String>> },
      FormatCells { range: String, format: CellFormat },
      CreateChart { data_range: String, chart_type: String },
  }
  ```
- [ ] Define agent response structure:
  ```rust
  pub struct AgentResponse {
      pub message_to_user: String,
      pub excel_commands: Vec<ExcelCommand>,
      pub confidence: f32,
  }
  ```

**Deliverable**: Type-safe Rust API for Excel operations

### 2.2 Tool Execution Engine
**Goal**: Parse LLM output and generate Excel commands

**Tasks**:
- [ ] Implement tool call parser (from LLM structured output)
- [ ] Build command validator (ensure ranges are valid, formulas parse, etc.)
- [ ] Implement command optimizer (batch operations, minimize syncs)
- [ ] Create error handling with user-friendly messages
- [ ] Add command preview/confirmation logic

**Deliverable**: Rust module that converts LLM tool calls → Excel commands

### 2.3 Excel Formula Generator
**Goal**: High-performance formula construction

**Tasks**:
- [ ] Implement formula builder utilities:
  - Range manipulation (A1:B10 → absolute/relative refs)
  - Formula validation
  - Common patterns (SUM, AVERAGE, VLOOKUP, etc.)
- [ ] Build formula suggestion engine
- [ ] Implement formula complexity analysis
- [ ] Add formula optimization (simplify redundant operations)

**Deliverable**: Rust library for robust formula generation

### 2.4 WASM Bindings
**Goal**: Expose Rust functions to JavaScript

**Tasks**:
- [ ] Use `wasm-bindgen` to expose main entry points:
  ```rust
  #[wasm_bindgen]
  pub fn process_agent_response(
      llm_output: &str,
      context: &str
  ) -> Result<String, JsValue> {
      // Parse LLM output
      // Generate commands
      // Return JSON
  }
  ```
- [ ] Implement JSON serialization for all public types
- [ ] Add JavaScript TypeScript definitions generation
- [ ] Optimize for size (use `wasm-opt`)

**Deliverable**: WASM module importable in Office.js add-in

---

## Phase 3: Integration (Week 5-6)

### 3.1 Agent Orchestration
**Goal**: Connect WebLLM, Rust core, and Office.js

**Tasks**:
- [ ] Implement agent loop in JavaScript:
  ```javascript
  async function handleUserMessage(message) {
      // 1. Get Excel context
      const context = await getExcelContext();
      
      // 2. Build prompt with tools
      const prompt = buildPrompt(message, context);
      
      // 3. LLM inference (WebLLM)
      const llmOutput = await engine.chat.completions.create({...});
      
      // 4. Process with Rust core
      const commands = await rustModule.process_agent_response(
          llmOutput, 
          JSON.stringify(context)
      );
      
      // 5. Execute in Excel (Office.js)
      await executeCommands(JSON.parse(commands));
  }
  ```
- [ ] Implement prompt engineering for tool calling
- [ ] Add system prompts for Excel-specific behavior
- [ ] Build conversation memory management

**Deliverable**: End-to-end message flow from user → LLM → Rust → Excel

### 3.2 Office.js Command Executor
**Goal**: Execute Rust-generated commands in Excel

**Tasks**:
- [ ] Implement command execution engine:
  ```javascript
  async function executeCommands(commands) {
      await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          
          for (const cmd of commands) {
              switch(cmd.type) {
                  case 'write_formula':
                      sheet.getRange(cmd.range).formulas = [[cmd.formula]];
                      break;
                  // ... more cases
              }
          }
          
          await context.sync();
      });
  }
  ```
- [ ] Add batch optimization (group operations)
- [ ] Implement undo/redo support
- [ ] Add error recovery and rollback

**Deliverable**: Reliable command execution system

### 3.3 Context Management
**Goal**: Efficiently pass Excel state to LLM

**Tasks**:
- [ ] Implement smart context extraction:
  - Limit data size (first N rows, summary stats)
  - Detect headers automatically
  - Include relevant formulas only
- [ ] Build context caching (avoid redundant reads)
- [ ] Implement incremental context updates
- [ ] Add context compression for large workbooks

**Deliverable**: Efficient context system that fits in LLM context window

---

## Phase 4: Optimization (Week 7-8)

### 4.1 Performance Tuning
**Goal**: Fast, responsive agent

**Tasks**:
- [ ] Profile WASM execution time
- [ ] Optimize hot paths in Rust (formula generation, parsing)
- [ ] Minimize JavaScript ↔ WASM boundary crossings
- [ ] Reduce WASM binary size:
  ```bash
  wasm-opt -Oz --strip-debug agent.wasm -o agent-opt.wasm
  ```
- [ ] Implement Web Worker for LLM inference (keep UI responsive)
- [ ] Add streaming responses (show LLM output as it generates)

**Deliverable**: <500ms latency for typical operations

### 4.2 Memory Management

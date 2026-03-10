# xlsx-review Makefile
# Build a single native binary using .NET 8 + Open XML SDK
#
# Usage:
#   make              # Build for current platform
#   make install      # Build + install to /usr/local/bin
#   make all          # Build for macOS ARM64, macOS x64, Linux x64, Linux ARM64
#   make docker       # Build Docker image
#   make smoke        # Run bundled example smoke tests
#   make test         # Run test against sample spreadsheet
#   make test-create  # Run create-mode smoke test
#   make corpus-download  # Download public XLSX regression corpus
#   make corpus-smoke     # Run a curated read smoke suite from the public corpus
#   make corpus-check     # Run read checks across the public corpus
#   make clean        # Remove build artifacts

BINARY_NAME  := xlsx-review
VERSION      := 1.1.0
BUILD_DIR    := build
INSTALL_DIR  := /usr/local/bin
LIMIT        ?= 50
LOCAL_RUNNER := ./scripts/run_local_release.sh

# Detect current platform
UNAME_S := $(shell uname -s)
UNAME_M := $(shell uname -m)

ifeq ($(UNAME_S),Darwin)
  ifeq ($(UNAME_M),arm64)
    CURRENT_RID := osx-arm64
  else
    CURRENT_RID := osx-x64
  endif
else
  CURRENT_RID := linux-x64
endif

# .NET publish flags
PUBLISH_FLAGS := -c Release \
  --self-contained \
  -p:PublishSingleFile=true \
  -p:IncludeNativeLibrariesForSelfExtract=true

.PHONY: build build-release install all docker smoke test test-dry test-create corpus-download corpus-smoke corpus-feature-smoke corpus-check corpus-check-fast clean help

## build: Build single-file binary for current platform
build:
	@echo "Building $(BINARY_NAME) for $(CURRENT_RID)..."
	@mkdir -p $(BUILD_DIR)
	dotnet publish $(PUBLISH_FLAGS) -r $(CURRENT_RID) -o $(BUILD_DIR)/$(CURRENT_RID)
	@cp $(BUILD_DIR)/$(CURRENT_RID)/$(BINARY_NAME) $(BUILD_DIR)/$(BINARY_NAME)
	@echo ""
	@echo "✅ Built: $(BUILD_DIR)/$(BINARY_NAME)"
	@ls -lh $(BUILD_DIR)/$(BINARY_NAME)

## build-release: Build the local release binary used for smoke/corpus checks
build-release:
	@echo "Building local release configuration..."
	@dotnet build -c Release >/dev/null
	@echo "✅ Ready: $(LOCAL_RUNNER)"

## install: Build and install to /usr/local/bin
install: build
	@echo "Installing to $(INSTALL_DIR)/$(BINARY_NAME)..."
	@cp $(BUILD_DIR)/$(BINARY_NAME) $(INSTALL_DIR)/$(BINARY_NAME)
	@chmod +x $(INSTALL_DIR)/$(BINARY_NAME)
	@echo "✅ Installed: $(INSTALL_DIR)/$(BINARY_NAME)"

## uninstall: Remove from /usr/local/bin
uninstall:
	@rm -f $(INSTALL_DIR)/$(BINARY_NAME)
	@echo "Removed $(INSTALL_DIR)/$(BINARY_NAME)"

## all: Build for all platforms (macOS ARM64, macOS x64, Linux x64, Linux ARM64)
all:
	@echo "Building for all platforms..."
	@mkdir -p $(BUILD_DIR)
	@for rid in osx-arm64 osx-x64 linux-x64 linux-arm64; do \
		echo ""; \
		echo "→ Building for $$rid..."; \
		dotnet publish $(PUBLISH_FLAGS) -r $$rid -o $(BUILD_DIR)/$$rid; \
		echo "  ✅ $(BUILD_DIR)/$$rid/$(BINARY_NAME)"; \
	done
	@echo ""
	@echo "All builds complete:"
	@ls -lh $(BUILD_DIR)/osx-arm64/$(BINARY_NAME) $(BUILD_DIR)/osx-x64/$(BINARY_NAME) $(BUILD_DIR)/linux-x64/$(BINARY_NAME) $(BUILD_DIR)/linux-arm64/$(BINARY_NAME) 2>/dev/null

## docker: Build Docker image
docker:
	docker build -t $(BINARY_NAME) .

## smoke: Run bundled read/diff/edit/create smoke tests
smoke: build-release
	@echo "Running bundled smoke tests..."
	@mkdir -p $(BUILD_DIR)
	@$(LOCAL_RUNNER) examples/test_old.xlsx --read --json > $(BUILD_DIR)/smoke-read.json
	@$(LOCAL_RUNNER) --diff examples/test_old.xlsx examples/test_new.xlsx --json > $(BUILD_DIR)/smoke-diff.json
	@$(LOCAL_RUNNER) examples/test_old.xlsx examples/sample-edits.json -o $(BUILD_DIR)/smoke-output.xlsx --json > $(BUILD_DIR)/smoke-edit.json
	@$(LOCAL_RUNNER) --create -o $(BUILD_DIR)/smoke-created.xlsx examples/sample-create.json --json > $(BUILD_DIR)/smoke-create.json
	@grep -q '"success": true' $(BUILD_DIR)/smoke-edit.json
	@grep -q '"success": true' $(BUILD_DIR)/smoke-create.json
	@$(LOCAL_RUNNER) $(BUILD_DIR)/smoke-created.xlsx --read --json > $(BUILD_DIR)/smoke-create-read.json
	@echo "✅ Smoke tests passed"
	@ls -lh $(BUILD_DIR)/smoke-output.xlsx $(BUILD_DIR)/smoke-created.xlsx

## test: Run test against a sample spreadsheet
test: build-release
	@echo "Running test..."
	@if [ ! -f examples/sample-edits.json ]; then \
		echo "Error: examples/sample-edits.json not found"; \
		exit 1; \
	fi
	@if [ ! -f "$(TEST_DOC)" ]; then \
		echo "Usage: make test TEST_DOC=/path/to/spreadsheet.xlsx"; \
		exit 1; \
	fi
	$(LOCAL_RUNNER) "$(TEST_DOC)" examples/sample-edits.json -o $(BUILD_DIR)/test_output.xlsx
	@echo ""
	@ls -lh $(BUILD_DIR)/test_output.xlsx

## test-dry: Dry-run test (no modifications)
test-dry: build-release
	@if [ -f "$(TEST_DOC)" ]; then \
		$(LOCAL_RUNNER) "$(TEST_DOC)" examples/sample-edits.json --dry-run; \
	else \
		echo "Usage: make test-dry TEST_DOC=/path/to/spreadsheet.xlsx"; \
	fi

## test-create: Exercise create mode with the sample manifest
test-create: build-release
	@echo "Testing --create mode..."
	@$(LOCAL_RUNNER) --create -o $(BUILD_DIR)/test_created.xlsx examples/sample-create.json --json > $(BUILD_DIR)/test-create.json
	@grep -q '"success": true' $(BUILD_DIR)/test-create.json
	@echo "Testing --create with template..."
	@$(LOCAL_RUNNER) --create --template examples/test_old.xlsx -o $(BUILD_DIR)/test_created_from_template.xlsx examples/sample-edits.json --json > $(BUILD_DIR)/test-create-template.json
	@grep -q '"success": true' $(BUILD_DIR)/test-create-template.json
	@echo "Testing --create dry-run..."
	@$(LOCAL_RUNNER) --create examples/sample-create.json --dry-run --json > $(BUILD_DIR)/test-create-dry.json
	@grep -q '"success": true' $(BUILD_DIR)/test-create-dry.json
	@ls -lh $(BUILD_DIR)/test_created.xlsx $(BUILD_DIR)/test_created_from_template.xlsx
	@echo "✅ Create tests passed"

## corpus-download: Download the public XLSX regression corpus
corpus-download:
	@./scripts/download_public_corpus.sh

## corpus-smoke: Run a curated read-mode smoke suite from the public corpus
corpus-smoke: build
	@./scripts/run_public_corpus_check.sh \
		--binary ./$(BUILD_DIR)/$(BINARY_NAME) \
		--mode read \
		--suite testdata/public-xlsx-corpus/suites/read-smoke.txt \
		--report-prefix read_smoke \
		--strict

## corpus-feature-smoke: Assert workbook/sheet metadata on representative corpus files
corpus-feature-smoke: build-release
	@./scripts/run_feature_smoke.sh --binary $(LOCAL_RUNNER)

## corpus-check: Run read checks across the full public XLSX corpus
corpus-check: build-release
	@./scripts/run_public_corpus_check.sh --binary $(LOCAL_RUNNER) --report-prefix read_full

## corpus-check-fast: Run a limited corpus check (override LIMIT=...)
corpus-check-fast: build-release
	@./scripts/run_public_corpus_check.sh --binary $(LOCAL_RUNNER) --limit $(LIMIT) --report-prefix read_fast

## clean: Remove build artifacts
clean:
	@rm -rf $(BUILD_DIR) bin obj
	@echo "Cleaned build artifacts"

## help: Show this help
help:
	@echo "xlsx-review $(VERSION) — Excel spreadsheet editing tool"
	@echo ""
	@echo "Targets:"
	@grep -E '^## ' Makefile | sed 's/## /  /' | column -t -s ':'
	@echo ""
	@echo "Examples:"
	@echo "  make                              # Build for $(CURRENT_RID)"
	@echo "  make install                      # Build + install to $(INSTALL_DIR)"
	@echo "  make all                          # Cross-compile all platforms"
	@echo "  make smoke                        # Run bundled example smoke tests"
	@echo "  make test-create                  # Exercise workbook creation mode"
	@echo "  make corpus-download              # Download public XLSX corpus"
	@echo "  make corpus-smoke                 # Run the curated corpus smoke suite"
	@echo "  make corpus-feature-smoke         # Assert workbook/sheet feature metadata"
	@echo "  make corpus-check                 # Validate the full public XLSX corpus"
	@echo "  make corpus-check-fast LIMIT=25   # Quick corpus subset check"
	@echo "  make test TEST_DOC=data.xlsx      # Run test"
	@echo "  make clean                        # Remove artifacts"

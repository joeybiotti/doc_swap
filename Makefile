# Load environment variables from .env
include .env
export

# Run the script with .env values
run:
	python main.py

# Run with dry-run mode (preview changes, don't save)
dry-run:
	python main.py --dry-run

# Create virtual environment and install dependencies
setup:
	python -m venv venv
	venv/bin/pip install -r requirements.txt

# Run tests (if you add a tests/ folder)
test:
	pytest tests/

# Clean output directory
clean:
	rm -rf $(OUTPUT_DIR)/*.docx
	rm -f $(OUTPUT_DIR)/log.txt

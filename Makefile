_python_pkg := google_sheets_wrapper

.PHONY: setup
setup:  ## Create virtual environment and install dependencies
	poetry install

.PHONY: test
test:  ## Run unit tests
	poetry run python -m unittest

.PHONY: lint
lint:  ## Run linting
	poetry run flake8 $(_python_pkg)
	poetry run mypy $(_python_pkg) --ignore-missing-imports
	poetry run isort -c $(_python_pkg)

.PHONY: reformat
reformat:  ## Reformat Python code using Black
	black -l 79 --skip-string-normalization $(_python_pkg)

.PHONY: help
help:
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-16s\033[0m %s\n", $$1, $$2}'

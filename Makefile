SCHEDULE = data/example/summit-tower-activities.csv
RISK_PARAMS = data/example/summit-tower-risk-params.csv
WORKBOOK = data/example/summit-tower-activities-monte-carlo.xlsx
ITERATIONS = 10000

all: gantt workbook simulate summary
	@echo ""
	@echo "Done. All outputs in outputs/"
	@echo "  Combined: outputs/summary.png"

gantt:
	python3 scripts/render_gantt.py --schedule $(SCHEDULE)

workbook:
	python3 scripts/monte_carlo.py --schedule $(SCHEDULE) --risk-params $(RISK_PARAMS) --init

simulate:
	python3 scripts/monte_carlo.py --schedule $(SCHEDULE) --workbook $(WORKBOOK) -n $(ITERATIONS)

summary:
	python3 scripts/combine_outputs.py

clean:
	rm -f outputs/gantt.png outputs/monte-carlo-*.png outputs/summary.png $(WORKBOOK)

.PHONY: all gantt workbook simulate summary clean

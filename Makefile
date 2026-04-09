SCHEDULE = data/example/summit-tower-activities.csv
RISK_PARAMS = data/example/summit-tower-risk-params.csv
WORKBOOK = data/example/summit-tower-activities-monte-carlo.xlsx
ITERATIONS = 10000

all: gantt workbook simulate
	@echo ""
	@echo "Done. All outputs in outputs/"

gantt:
	python3 scripts/render_gantt.py --schedule $(SCHEDULE)

workbook:
	python3 scripts/monte_carlo.py --schedule $(SCHEDULE) --risk-params $(RISK_PARAMS) --init

simulate:
	python3 scripts/monte_carlo.py --schedule $(SCHEDULE) --workbook $(WORKBOOK) -n $(ITERATIONS)

clean:
	rm -f outputs/gantt.png outputs/monte-carlo-*.png $(WORKBOOK)

.PHONY: all gantt workbook simulate clean

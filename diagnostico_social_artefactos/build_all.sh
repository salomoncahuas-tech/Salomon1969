set -e
sh run_hua.sh
sh run_mor.sh
python3 mor_html.py
python3 aya_html.py
python3 term_apply.py

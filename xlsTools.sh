#!/bin/bash

python3 scripts/xlsTools.py --input_dir ./xls \
    --client_type json \
    --client_output_dir ./output/client \
    --server_type lua \
    --server_output_dir ./output/server \
    --exclude_files .git .svn $@

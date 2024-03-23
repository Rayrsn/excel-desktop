#!/usr/bin/env bash
ls *py | entr bash -c "python ./pw_query.py; "
  # echo \"=====================\"
  # echo -e \"\n\n\"

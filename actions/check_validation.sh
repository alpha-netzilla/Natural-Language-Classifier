#!/bin/sh -l


grep -r -E "[\"\'][0-9a-zA-Z\-]{44,}[\"\']" ./*.gs > /dev/null

if [ $? -eq 0 ]; then
  echo "Include invalie characters"
  exit 1
else
  echo "Successful validation"
  exit 0
fi

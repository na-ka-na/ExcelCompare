#!/usr/bin/env sh

if [ -L $0 ];then
  dir=`readlink -f $0|xargs dirname`
else
  dir=`dirname $0`
fi
java -ea -cp "$dir/dist/*" com.ka.spreadsheet.diff.SpreadSheetDiffer "$@"

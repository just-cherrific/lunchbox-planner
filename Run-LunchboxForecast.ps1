$ErrorActionPreference = 'Stop'
$node = 'C:\Users\Linda\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe'
$env:NODE_PATH = 'C:\Users\Linda\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules'
& $node 'C:\trial3\lunchbox_forecast.mjs'
exit $LASTEXITCODE

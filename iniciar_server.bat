@echo off
title Servidor Experimental - TheTom TurboQuant
color 0D

echo ===================================================
echo Iniciando
echo ===================================================
echo.

llama-server -m "C:\Users\arthurvfr\OneDrive - CGU\Documentos\modelos_llamacpp\Qwen3-Coder-30B-A3B-Instruct-UD-Q8_K_XL.gguf" ^
  --port 1234 ^
  --n-gpu-layers 999 ^
  --n-cpu-moe 36 ^
  --no-mmap ^
  --mlock ^
  --cache-type-k q4_0 --cache-type-v q4_0 ^
  --repeat_penalty 1.1 ^
  --mirostat 0 ^
  --temp 0 ^
  -c 8192 ^
  --cont_batching


echo.
echo O servidor foi encerrado.
pause
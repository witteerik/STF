// License
// This project is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
// Commercial use requires a separate commercial license.See the `LICENSE` file for details.
// 
// Copyright (c) 2025 Erik Witte

#pragma once

#include "libstfdspandroid/export.h"

extern "C" STFDSPANDROID_API void copyToDouble(float* sourceArray, int size, double* targetArray);

extern "C" STFDSPANDROID_API void copyToFloat(double* sourceArray, int size, float* targetArray);

extern "C" STFDSPANDROID_API void createInterleavedArray(float* concatenatedArrays, int channelCount, int channelLength, float* interleavedArray, bool applyGain, float* channelGainFactor);

extern "C" STFDSPANDROID_API void deinterleaveArray(float* interleavedArray, int channelCount, int channelLength, float* concatenatedArrays);

extern "C" STFDSPANDROID_API int multiplyDoubleArray(double* values, int size, double factor);

extern "C" STFDSPANDROID_API int multiplyDoubleArraySection(double* values, int arraySize, double factor, int startIndex, int sectionLength);

extern "C" STFDSPANDROID_API int multiplyFloatArray(float* values, int size, float factor);

extern "C" STFDSPANDROID_API int multiplyFloatArraySection(float* values, int arraySize, float factor, int startIndex, int sectionLength);

extern "C" STFDSPANDROID_API double calculateFloatSumOfSquare(float* values, int arraySize, int startIndex, int sectionLength);

extern "C" STFDSPANDROID_API void addTwoFloatArrays(
	float* array1,
	float* array2,
	int size);

extern "C" STFDSPANDROID_API void fft_complex(
	double* real,
	double* imag,
	int size,
	double* pccos,
	double* pcsin,
	int direction,
	bool reorder,
	bool scaleForwardTransform);

extern "C" STFDSPANDROID_API void complexMultiplication(double* real1, double* imag1, double* real2, double* imag2, int size);

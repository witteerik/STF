// License
// This project is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
// Commercial use requires a separate commercial license.See the `LICENSE` file for details.
// 
// Copyright (c) 2025 Erik Witte

#include "libstfdspandroid/dsp.h"

#include <algorithm>
#include <cfloat>
#include <cmath>

void copyToDouble(float* sourceArray, int size, double* targetArray)
{
	for (int i = 0; i < size; i++)
	{
		targetArray[i] = sourceArray[i];
	}
}

void copyToFloat(double* sourceArray, int size, float* targetArray)
{
	const double min = -FLT_MAX;
	const double max = FLT_MAX;

	for (int i = 0; i < size; i++)
	{
		const double limitedValue = std::clamp(sourceArray[i], min, max);
		targetArray[i] = static_cast<float>(limitedValue);
	}
}

void createInterleavedArray(float* concatenatedArrays, int channelCount, int channelLength, float* interleavedArray, bool applyGain, float* channelGainFactor)
{
	int targetIndex = 0;
	if (applyGain)
	{
		for (int sampleIndex = 0; sampleIndex < channelLength; sampleIndex++)
		{
			for (int channelIndex = 0; channelIndex < channelCount; channelIndex++)
			{
				interleavedArray[targetIndex] = channelGainFactor[channelIndex] * concatenatedArrays[channelIndex * channelLength + sampleIndex];
				targetIndex += 1;
			}
		}
	}
	else
	{
		for (int sampleIndex = 0; sampleIndex < channelLength; sampleIndex++)
		{
			for (int channelIndex = 0; channelIndex < channelCount; channelIndex++)
			{
				interleavedArray[targetIndex] = concatenatedArrays[channelIndex * channelLength + sampleIndex];
				targetIndex += 1;
			}
		}
	}
}

void deinterleaveArray(float* interleavedArray, int channelCount, int channelLength, float* concatenatedArrays)
{
	int targetIndex = 0;
	for (int sampleIndex = 0; sampleIndex < channelLength; sampleIndex++)
	{
		for (int channelIndex = 0; channelIndex < channelCount; channelIndex++)
		{
			concatenatedArrays[channelIndex * channelLength + sampleIndex] = interleavedArray[targetIndex];
			targetIndex += 1;
		}
	}
}

int multiplyDoubleArray(double* values, int size, double factor)
{
	return multiplyDoubleArraySection(values, size, factor, 0, size);
}

int multiplyDoubleArraySection(double* values, int arraySize, double factor, int startIndex, int sectionLength)
{
	int clippedCount = 0;
	const double min = -DBL_MAX;
	const double max = DBL_MAX;

	if (arraySize < 1)
	{
		return clippedCount;
	}

	startIndex = std::clamp(startIndex, 0, arraySize - 1);
	sectionLength = std::clamp(sectionLength, 0, arraySize - startIndex);

	for (int i = startIndex; i < startIndex + sectionLength; i++)
	{
		const double newValue = values[i] * factor;
		const double limitedValue = std::clamp(newValue, min, max);
		if (newValue != limitedValue)
		{
			clippedCount++;
		}

		values[i] = limitedValue;
	}

	return clippedCount;
}

int multiplyFloatArray(float* values, int size, float factor)
{
	return multiplyFloatArraySection(values, size, factor, 0, size);
}

int multiplyFloatArraySection(float* values, int arraySize, float factor, int startIndex, int sectionLength)
{
	int clippedCount = 0;
	const double min = -FLT_MAX;
	const double max = FLT_MAX;

	if (arraySize < 1)
	{
		return clippedCount;
	}

	startIndex = std::clamp(startIndex, 0, arraySize - 1);
	sectionLength = std::clamp(sectionLength, 0, arraySize - startIndex);

	for (int i = startIndex; i < startIndex + sectionLength; i++)
	{
		const double newValue = values[i] * factor;
		const double limitedValue = std::clamp(newValue, min, max);
		if (newValue != limitedValue)
		{
			clippedCount++;
		}

		values[i] = static_cast<float>(limitedValue);
	}

	return clippedCount;
}

double calculateFloatSumOfSquare(float* values, int arraySize, int startIndex, int sectionLength)
{
	startIndex = std::clamp(startIndex, 0, arraySize - 1);
	sectionLength = std::clamp(sectionLength, 0, arraySize - startIndex);

	double sumOfSquare = 0;
	for (int i = startIndex; i < startIndex + sectionLength; i++)
	{
		sumOfSquare += values[i] * values[i];
	}

	return sumOfSquare;
}

void addTwoFloatArrays(float* array1, float* array2, int size)
{
	for (int i = 0; i < size; i++)
	{
		array1[i] += array2[i];
	}
}

void fft_complex(double* real, double* imag, int size, double* pccos, double* pcsin, int direction, bool reorder, bool scaleForwardTransform)
{
	int exponentSign = 0;
	if (direction == 1)
	{
		exponentSign = -1;
	}
	else
	{
		exponentSign = 1;
	}

	if (reorder)
	{
		double tempX = 0;
		double tempY = 0;

		int j = 0;
		int m = 0;

		for (int index = 0; index < size - 1; ++index)
		{
			if (index < j)
			{
				tempX = real[index];
				real[index] = real[j];
				real[j] = tempX;

				tempY = imag[index];
				imag[index] = imag[j];
				imag[j] = tempY;
			}

			m = size;
			do
			{
				m >>= 1;
				j ^= m;
			} while ((j & m) == 0);
		}
	}

	const int halfSize = size / 2;

	double aiX = 0;
	double aiY = 0;
	double real1 = 0;
	double imaginary1 = 0;
	double real2 = 0;
	double imaginary2 = 0;
	double tempReal1 = 0;
	int lookupIndex = 0;
	double wX = 0;
	double wY = 0;
	int stepSize = 0;
	int i = 0;
	int levelSize = 1;

	while (levelSize < size)
	{
		stepSize = levelSize << 1;

		for (int k = 0; k < levelSize; ++k)
		{
			lookupIndex = static_cast<int>(static_cast<double>(halfSize) * (static_cast<double>(k) / static_cast<double>(levelSize)));

			wX = pccos[lookupIndex];
			wY = pcsin[lookupIndex];

			i = k;
			while (i < size - 1)
			{
				aiX = real[i];
				aiY = imag[i];

				real1 = wX;
				imaginary1 = wY;
				real2 = real[i + levelSize];
				imaginary2 = imag[i + levelSize];

				tempReal1 = real1;
				real1 = tempReal1 * real2 - imaginary1 * imaginary2;
				imaginary1 = tempReal1 * imaginary2 + imaginary1 * real2;

				real[i] = aiX + real1;
				imag[i] = aiY + imaginary1;

				real[i + levelSize] = aiX - real1;
				imag[i + levelSize] = aiY - imaginary1;

				i += stepSize;
			}
		}

		levelSize *= 2;
	}

	if (direction == 1 && scaleForwardTransform)
	{
		const double scalingFactor = 1.0 / size;
		for (int index = 0; index < size; ++index)
		{
			real[index] *= scalingFactor;
			imag[index] *= scalingFactor;
		}
	}
}

void complexMultiplication(double* real1, double* imag1, double* real2, double* imag2, int size)
{
	for (int i = 0; i < size; i++)
	{
		const double tempValue = real1[i];
		real1[i] = tempValue * real2[i] - imag1[i] * imag2[i];
		imag1[i] = tempValue * imag2[i] + imag1[i] * real2[i];
	}
}

#pragma once

#if defined(_WIN32)
	#if defined(STFDSPANDROID_EXPORTS)
		#define STFDSPANDROID_API __declspec(dllexport)
	#else
		#define STFDSPANDROID_API __declspec(dllimport)
	#endif
#else
	#define STFDSPANDROID_API __attribute__((visibility("default")))
#endif

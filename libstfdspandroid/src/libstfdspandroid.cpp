#include "libstfdspandroid/libstfdspandroid.h"

#include <android/log.h>

namespace
{
	constexpr char LogTag[] = "libstfdspandroid";
}

extern "C" STFDSPANDROID_API const char* stfdspandroid_getPlatformABI()
{
	const char* abi = "unknown";
#if defined(__aarch64__)
	abi = "arm64-v8a";
#elif defined(__arm__)
	#if defined(__ARM_ARCH_7A__)
		#if defined(__ARM_NEON__)
			abi = "armeabi-v7a/NEON";
		#else
			abi = "armeabi-v7a";
		#endif
	#else
		abi = "armeabi";
	#endif
#elif defined(__i386__)
	abi = "x86";
#elif defined(__x86_64__)
	abi = "x86_64";
#endif

	__android_log_print(ANDROID_LOG_INFO, LogTag, "This dynamic shared library is compiled with ABI: %s", abi);
	return abi;
}

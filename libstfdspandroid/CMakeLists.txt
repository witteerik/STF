cmake_minimum_required(VERSION 3.22.1)

project(libstfdspandroid VERSION 1.0.0 LANGUAGES C CXX)

set(CMAKE_C_STANDARD 11)
set(CMAKE_C_STANDARD_REQUIRED ON)
set(CMAKE_C_EXTENSIONS OFF)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

add_library(libstfdspandroid SHARED
	src/dsp.cpp
	src/libstfdspandroid.cpp
)

target_include_directories(libstfdspandroid
	PUBLIC
		${CMAKE_CURRENT_SOURCE_DIR}/include
)

find_library(android-log-lib log)

target_link_libraries(libstfdspandroid
	PRIVATE
		${android-log-lib}
)

target_compile_definitions(libstfdspandroid PRIVATE STFDSPANDROID_EXPORTS)

target_compile_options(libstfdspandroid PRIVATE
	$<$<COMPILE_LANGUAGE:CXX>:-Wall -Wextra -Wpedantic>
)

set_target_properties(libstfdspandroid PROPERTIES
	OUTPUT_NAME "libostfdspandroid"
	POSITION_INDEPENDENT_CODE ON
)

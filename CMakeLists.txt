
cmake_minimum_required(VERSION 3.0)
project(SpdlogExamples)

find_package(Threads)

include_directories(/home/mickael/example/include)

target_link_libraries( ${CMAKE_THREAD_LIBS_INIT})


set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -std=gnu++14")


set(SOURCE_FILES main.cpp include/spdlog/log.h include/GlobalConst.h)
add_executable(exe ${SOURCE_FILES})

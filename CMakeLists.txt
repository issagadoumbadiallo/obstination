cmake_minimum_required(VERSION 3.5)
project(HelloClion)

set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -std=gnu++14")

#include_directories(C:/include/spdlog)

set(SOURCE_FILES main.cpp include/spdlog/log.h include/GlobalConst.h)
add_executable(HelloClion ${SOURCE_FILES})


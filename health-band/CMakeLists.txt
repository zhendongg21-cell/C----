# 顶层 CMake：注册 hal / driver / service / app 四组件
# 「目录即架构」：用 target_link_libraries 强制「上层依赖下层」，编译期阻止反向依赖。
# P0 阶段为目录骨架；各组件内 CMakeLists.txt 在 P1/P2 随代码补齐。
cmake_minimum_required(VERSION 3.16)

# EXTRA_COMPONENT_DIRS 指向 main/ 下的分层组件目录
set(EXTRA_COMPONENT_DIRS
    "main/hal"
    "main/driver"
    "main/service"
    "main/app"
    "main/ui"
    "main/board"
)

include($ENV{IDF_PATH}/tools/cmake/project.cmake)
project(health-band)

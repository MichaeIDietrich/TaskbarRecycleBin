#pragma once

#ifdef DEBUG
#define LOGGING
#endif

#ifdef LOGGING

#include <fstream>
#include <time.h>
#include <iomanip>

#define LOGGING_PATH "RecycleBin.log" // store it into module path

#define LOG(s) { time_t current_time = time(nullptr); \
                 tm time_info; \
                 localtime_s(&time_info, &current_time); \
                 std::ofstream _log(LOGGING_PATH, std::ofstream::out | std::ofstream::app); \
                 _log << std::put_time(&time_info, "%Y-%m-%d %X "); \
                 _log << s << "\n"; \
                 _log.close(); \
               }

#define LOG2(s, s2) { time_t current_time = time(nullptr); \
                      tm time_info; \
                      localtime_s(&time_info, &current_time); \
                      std::ofstream _log(LOGGING_PATH, std::ofstream::out | std::ofstream::app); \
                      _log << std::put_time(&time_info, "%Y-%m-%d %X "); \
                      _log << s << ": " << s2 << "\n"; \
                      _log.close(); }

#else
#define LOG(s);
#define LOG2(s, s2);
#endif
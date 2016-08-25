#include "spdlog.h"
//#include "logger.h"
//#include "tweakme.h"
//#include "common.h"

class LivingObj {
    std::shared_ptr<spdlog::logger> file;
    size_t q_size = 1048576;//queue size must be power of 2
public:
    LivingObj(){
        spdlog::set_async_mode(q_size);
        file = spdlog::daily_logger_st( "EngineLog", "log.txt");
    };

    ~LivingObj(){
        spdlog::drop_all();
    };

    template<typename T> void operator<<( const T* msg)
    {
        file->info(msg);
    }
/*    void i(const char* msg){
        file->info(msg);
    };*/

};


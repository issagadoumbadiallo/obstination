//#include <iostream>
//#include "include/GlobalConst.h"
//#include <vector>
//#include <chrono>
//#include <time.h>
//#include <thread>
#include <iostream>
#include <iomanip>
#include <chrono>
#include <ctime>
#include <thread>
void f()
{
    volatile double d = 0;
    for(int n=0; n<10000; ++n)
        for(int m=0; m<10000; ++m)
            d += d*n*m;
}

/*void tt(){
    l<<"Trying tt voidFunc";
}*/
//
/*int main() {
    l << "Starting Main";
    d.restart();

    std::thread t1(f);
    std::thread t2(f); // f() is called on two threads
    std::thread t3(f); // f() is called on two threads
    t1.join();
    t2.join();
    t3.join();

    d.end();
    d.show_all();

    return 0;
}*/

int main()
{
    std::clock_t c_start = std::clock();
    auto t_start = std::chrono::high_resolution_clock::now();
    std::thread t1(f);
    std::thread t2(f); // f() is called on two threads
    std::thread t3(f); // f() is called on two threads

    t1.join();
    t2.join();
    t3.join();

    std::clock_t c_end = std::clock();
    auto t_end = std::chrono::high_resolution_clock::now();

    std::cout << std::fixed << std::setprecision(2) << "CPU time used: "
              << 1000.0 * (c_end-c_start) / CLOCKS_PER_SEC << " ms\n"
              << "Wall clock time passed: "
              << std::chrono::duration<double, std::milli>(t_end-t_start).count()
              << " ms\n";
}

/* int main() {
    LivingObj l= LivingObj();
    l << "hello" ;

    l << "1"; //"Test";

    return 0;
}*/

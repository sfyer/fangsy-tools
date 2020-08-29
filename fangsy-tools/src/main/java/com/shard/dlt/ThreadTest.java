package com.shard.dlt;

import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

/**
 * @Author : fsy
 * @Date: 2020-08-29 09:53
 */
public class ThreadTest {

    public static void main(String[] args) {
        ExecutorService executorService = new ThreadPoolExecutor(10, 10
            , 90, TimeUnit.MINUTES, new LinkedBlockingQueue<Runnable>(800)
            , new ThreadPoolExecutor.CallerRunsPolicy());

        /**
         *acc : 获取调用上下文
         * 	corePoolSize: 核心线程数量，可以类比正式员工数量，常驻线程数量。
         * 	maximumPoolSize: 最大的线程数量，公司最多雇佣员工数量。常驻+临时线程数量。
         * 	workQueue：多余任务等待队列，再多的人都处理不过来了，需要等着，在这个地方等。
         * 	keepAliveTime：非核心线程空闲时间，就是外包人员等了多久，如果还没有活干，解雇了。
         * 	threadFactory: 创建线程的工厂，在这个地方可以统一处理创建的线程的属性。
         * 	每个公司对员工的要求不一样，恩，在这里设置员工的属性。
         * 	handler：线程池拒绝策略，什么意思呢?就是当任务实在是太多，人也不够，需求池也排满了，
         * 	还有任务咋办?默认是不处理，抛出异常告诉任务提交者，我这忙不过来了。
         */
        ExecutorService executorService1 = Executors.newFixedThreadPool(10);

    }

}

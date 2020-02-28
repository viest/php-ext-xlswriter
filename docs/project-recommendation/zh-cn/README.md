## 1. [EasySwoole](https://www.easyswoole.com/)

```
  ______                          _____                              _        
 |  ____|                        / ____|                            | |       
 | |__      __ _   ___   _   _  | (___   __      __   ___     ___   | |   ___ 
 |  __|    / _` | / __| | | | |  \___ \  \ \ /\ / /  / _ \   / _ \  | |  / _ \
 | |____  | (_| | \__ \ | |_| |  ____) |  \ V  V /  | (_) | | (_) | | | |  __/
 |______|  \__,_| |___/  \__, | |_____/    \_/\_/    \___/   \___/  |_|  \___|
                          __/ |                                               
                         |___/ 
```

EasySwoole 是一款基于Swoole Server 开发的常驻内存型的分布式PHP框架，专为API而生，摆脱传统PHP运行模式在进程唤起和文件加载上带来的性能损失。 EasySwoole 高度封装了 Swoole Server 而依旧维持 Swoole Server 原有特性，支持同时混合监听HTTP、自定义TCP、UDP协议，让开发者以最低的学习成本和精力编写出多进程，可异步，高可用的应用服务。在开发上，我们为您准备了以下常用组件：

 - http 服务服务器
 - 协程ORM(类似Tp Orm)
 - 图片验证码
 - validate验证器
 - 协程模板渲染引擎
 - jwt组件
 - 协程TCP、UDP、WEB_SOCKET 服务端
 - 协程redis连接池
 - 协程mysql 连接池
 - 协程Memcached客户端
 - 协程通用链接池
 - 协程kafka客户端
 - 分布式跨平台RPC组件
 - 协程consul客户端
 - 协程apollo配置中心
 - 协程Actor
 - 协程Smtp客户端
 - 协程微信公众号与小程序SDK
 - 协程协程版微信、支付宝支付SDK
 - 协程elasticsearch客户端
 - 协程HTTP客户端组件
 - 协程上下文管理
 - IOC、协程上下文管理器
 - snowflake Id生成器
 - crontab 秒级任务
 - 自定义进程
 - 自定义消息队列
 - Tracker链路跟踪
 - Atomic限流器
 - fast-cache组件
 - policy权限组件
 
 ---
 
 ## 特性
 
 - 强大的 TCP/UDP Server 框架，多线程，EventLoop，事件驱动，异步，Worker进程组，Task异步任务，毫秒定时器，SSL/TLS隧道加密
 - EventLoop API，让用户可以直接操作底层的事件循环，将socket，stream，管道等Linux文件加入到事件循环中
 - 定时器、协程对象池、HTTP\SOCK控制器、分布式微服务、RPC支持
 
 ---
 
 ## 优势
 
 - 简单易用开发效率高
 - 并发百万TCP连接
 - TCP/UDP/UnixSock
 - 支持异步/同步/协程
 - 支持多进程/多线程
 - CPU亲和性/守护进程
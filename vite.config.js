import { defineConfig } from 'vite' // 使用 defineConfig 帮手函数，这样不用 jsdoc 注解也可以获取类型提示

export default ({ command, mode }) => {
    return defineConfig({
        server: {
            port: 9000, // 本地服务端口
            open: true, //在服务器启动时自动在浏览器中打开应用程序。当此值为字符串时，会被用作 URL 的路径名。
            strictPort: true, // 设为 true 时若端口已被占用则会直接退出，而不是尝试下一个可用端口
        },
        // 打包配置
        build: {
            target: 'modules',
            outDir: 'dist', //指定输出路径
            assetsDir: 'assets', // 指定生成静态资源的存放路径
            minify: 'terser', // 混淆器，terser构建后文件体积更小
            chunkSizeWarningLimit: 500000
        }
    })
}


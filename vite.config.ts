import {defineConfig} from 'vite';

export default defineConfig({
    build: {
        modulePreload: {
           polyfill: false
        },
        rollupOptions: {
            input: {
                sidebar: 'sidebar.html',
                about: 'about.html'
            }
        }
    },
    server: {
        open: 'sidebar.html'
    }
})
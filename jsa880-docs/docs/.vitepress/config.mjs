import { defineConfig } from 'vitepress'

export default defineConfig({
  title: 'JSA880',
  description: '郑广学JSA880快速开发框架 - WPS Office JavaScript API',

  lang: 'zh-CN',
  cleanUrls: true,
  lastUpdated: true,

  ignoreDeadLinks: true,

  themeConfig: {
    logo: '/img/logo.svg',
    siteTitle: 'JSA880',

    nav: [
      { text: '首页', link: '/' },
      { text: 'API', link: '/api/' },
      { text: '指南', link: '/guide/' },
      {
        text: '语言',
        items: [
          { text: '简体中文', link: '/zh/' },
          { text: 'English', link: '/en/' }
        ]
      }
    ],

    sidebar: {
      '/api/': [
        {
          text: 'API 参考',
          items: [
            { text: '概述', link: '/api/' },
            { text: 'JSA 全局函数', link: '/api/global-functions' },
            { text: 'Array2D 类', link: '/api/array2d-class' },
            { text: 'RngUtils 类', link: '/api/rngutils-class' },
            { text: 'ShtUtils 类', link: '/api/shtutils-class' }
          ]
        }
      ],
      '/guide/': [
        {
          text: '指南',
          items: [
            { text: '概述', link: '/guide/' },
            { text: '快速开始', link: '/guide/getting-started' },
            { text: 'Lambda 表达式', link: '/guide/lambda' },
            { text: '链式调用', link: '/guide/chaining' },
            { text: 'superPivot 透视表', link: '/guide/super-pivot' }
          ]
        }
      ]
    },

    socialLinks: [
      { icon: 'github', link: 'https://github.com/your-repo/jsa880' }
    ],

    footer: {
      message: '基于郑广学 JSA880 框架构建',
      copyright: 'Copyright © 2024-2026'
    },

    search: {
      provider: 'local'
    }
  },

  markdown: {
    lineNumbers: true,
    theme: {
      light: 'github-light',
      dark: 'github-dark'
    }
  }
})
<!-- このコードを header.php の <?php wp_head(); ?> の前に追加してください -->

<?php
// SEO設定：特定の固定ページにタイトルとメタディスクリプションを追加
if (is_page('career')) : ?>
    <title>採用情報・求人｜プラントエンジニア募集中｜コモンウェルスエンジニアーズ</title>
    <meta name="description" content="コモンウェルスエンジニアーズでは、プラントエンジニア・設計技術者を募集しています。未経験歓迎、充実した研修制度あり。国内外の大型プロジェクトに携わるチャンスです。">
    <link rel="canonical" href="https://cectokyo.com/career/">
<?php elseif (is_page('contact')) : ?>
    <title>お問い合わせ・ご相談｜プラント設計のコモンウェルスエンジニアーズ</title>
    <meta name="description" content="発電所・化学プラント設計に関するお問い合わせ・お見積りはこちら。50年以上の実績を持つコモンウェルスエンジニアーズが、国内外のプロジェクトをサポートします。お気軽にご相談ください。">
    <link rel="canonical" href="https://cectokyo.com/contact/">
<?php elseif (is_page('entry')) : ?>
    <title>採用エントリーフォーム｜中途・新卒採用｜コモンウェルスエンジニアーズ</title>
    <meta name="description" content="コモンウェルスエンジニアーズの採用エントリーフォームです。新卒・中途どちらも歓迎。書類選考から面接まで最短2週間。エンジニアとしてのキャリアをスタートしませんか。">
    <link rel="canonical" href="https://cectokyo.com/entry/">
<?php endif; ?>









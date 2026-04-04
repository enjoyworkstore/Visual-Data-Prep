import React from 'react';

// --- Icons ---
const IconSvg = ({ children, className = "w-6 h-6" }: any) => (
  <svg viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round" className={className}>{children}</svg>
);
const Icons = {
  Diamond: <IconSvg><polygon points="12 2 22 8.5 22 15.5 12 22 2 15.5 2 8.5 12 2"/></IconSvg>,
  Download: <IconSvg><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></IconSvg>,
  Github: <IconSvg><path d="M9 19c-5 1.5-5-2.5-7-3m14 6v-3.87a3.37 3.37 0 0 0-.94-2.61c3.14-.35 6.44-1.54 6.44-7A5.44 5.44 0 0 0 20 4.77 5.07 5.07 0 0 0 19.91 1S18.73.65 16 2.48a13.38 13.38 0 0 0-7 0C6.27.65 5.09 1 5.09 1A5.07 5.07 0 0 0 5 4.77a5.44 5.44 0 0 0-1.5 3.78c0 5.42 3.3 6.61 6.44 7A3.37 3.37 0 0 0 9 18.13V22"/></IconSvg>,
  Database: <IconSvg><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/></IconSvg>,
  Zap: <IconSvg><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></IconSvg>,
  Layout: <IconSvg><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="9" y1="21" x2="9" y2="9"/></IconSvg>,
  Code: <IconSvg><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></IconSvg>,
  ArrowRight: <IconSvg><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></IconSvg>
};

export default function LandingPage() {
  return (
    <div className="min-h-screen bg-[#0B0F19] text-gray-200 font-sans selection:bg-blue-500/30">
      
      {/* Navigation */}
      <nav className="border-b border-gray-800 bg-[#0B0F19]/80 backdrop-blur-md sticky top-0 z-50">
        <div className="max-w-6xl mx-auto px-6 py-4 flex justify-between items-center">
          <div className="flex items-center gap-3">
            <span className="text-blue-500">{Icons.Diamond}</span>
            <span className="text-lg font-bold tracking-[0.3em] text-white">VISUAL DATA PREP</span>
          </div>
          <div className="flex items-center gap-6">
            <a href="#features" className="text-sm font-semibold hover:text-white transition-colors hidden md:block">Features</a>
            <a href="#how-it-works" className="text-sm font-semibold hover:text-white transition-colors hidden md:block">How it Works</a>
            <a href="https://github.com/your-username/visual-data-prep" target="_blank" rel="noreferrer" className="text-gray-400 hover:text-white transition-colors">
              {Icons.Github}
            </a>
            <a href="#download" className="bg-blue-600 hover:bg-blue-500 text-white text-sm font-bold px-5 py-2.5 rounded-lg tracking-widest transition-all shadow-[0_0_15px_rgba(37,99,235,0.5)] flex items-center gap-2">
              {Icons.Download} DOWNLOAD
            </a>
          </div>
        </div>
      </nav>

      {/* Hero Section */}
      <section className="relative pt-32 pb-20 overflow-hidden">
        {/* Decorative background glow */}
        <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[800px] h-[400px] bg-blue-600/20 blur-[120px] rounded-full pointer-events-none"></div>
        
        <div className="max-w-6xl mx-auto px-6 relative z-10 text-center">
          <h1 className="text-5xl md:text-7xl font-extrabold text-white tracking-tight leading-tight mb-6">
            No-Code Data Reshaping <br className="hidden md:block" />
            <span className="text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-fuchsia-400">
              & Visual SQL Building
            </span>
          </h1>
          <p className="text-lg md:text-xl text-gray-400 max-w-3xl mx-auto mb-10 leading-relaxed">
            ドラッグ＆ドロップの直感的な操作で、CSVやWeb APIなどのデータを自由自在に結合・整形。複雑なデータ処理パイプラインやSQLクエリを誰でも手軽に構築できるデスクトップツールです。
          </p>
          <div className="flex flex-col sm:flex-row justify-center items-center gap-4">
            <a href="#download" className="bg-blue-600 hover:bg-blue-500 text-white text-sm font-bold px-8 py-4 rounded-xl tracking-widest transition-all shadow-[0_0_20px_rgba(37,99,235,0.4)] flex items-center gap-2 w-full sm:w-auto justify-center">
              {Icons.Download} WINDOWS版を無料ダウンロード
            </a>
            <a href="#how-it-works" className="bg-gray-800 hover:bg-gray-700 text-white text-sm font-bold px-8 py-4 rounded-xl tracking-widest transition-all flex items-center gap-2 border border-gray-700 w-full sm:w-auto justify-center">
              使い方を見る {Icons.ArrowRight}
            </a>
          </div>

          {/* App Mockup Placeholder */}
          <div className="mt-20 relative mx-auto max-w-5xl">
            <div className="rounded-2xl border border-gray-800 bg-[#1e1e1e] p-2 shadow-2xl relative">
              <div className="absolute inset-0 bg-gradient-to-t from-[#0B0F19] via-transparent to-transparent z-20 rounded-2xl pointer-events-none"></div>
              {/* ここにアプリのスクショ画像を配置します */}
              <img 
                src="https://placehold.co/1200x675/1e1e1e/38bdf8?text=App+Screenshot+Here" 
                alt="Visual Data Prep Screenshot" 
                className="w-full rounded-xl opacity-90 border border-gray-800"
              />
            </div>
          </div>
        </div>
      </section>

      {/* Features Section */}
      <section id="features" className="py-24 bg-[#0F1523]">
        <div className="max-w-6xl mx-auto px-6">
          <div className="text-center mb-16">
            <h2 className="text-sm font-bold tracking-[0.3em] text-blue-500 mb-2 uppercase">Features</h2>
            <h3 className="text-3xl md:text-4xl font-bold text-white">圧倒的に手軽なデータプレパレーション</h3>
          </div>
          
          <div className="grid md:grid-cols-2 lg:grid-cols-4 gap-8">
            {[
              { i: Icons.Layout, t: "直感的なノーコードUI", d: "ノードをキャンバスに配置して線で繋ぐだけ。プログラミングの知識は一切不要です。" },
              { i: Icons.Database, t: "多彩なデータソース", d: "ローカルのCSV/Excelだけでなく、フォルダ自動監視やWeb APIからの直接取得に対応。" },
              { i: Icons.Zap, t: "強力なクレンジング", d: "VLOOKUP的な結合、文字列抽出、ゼロ埋め、四則演算まで豊富な変換ノードを搭載。" },
              { i: Icons.Code, t: "SQLの相互変換", d: "作成したフローからSELECT文を自動生成。逆にSQLからノードを自動配置することも可能。" }
            ].map((f, idx) => (
              <div key={idx} className="bg-[#182032] border border-gray-800 p-6 rounded-2xl hover:border-blue-500/50 transition-colors">
                <div className="w-12 h-12 bg-blue-500/10 rounded-xl flex items-center justify-center text-blue-400 mb-4 border border-blue-500/20">
                  {f.i}
                </div>
                <h4 className="text-lg font-bold text-white mb-2">{f.t}</h4>
                <p className="text-sm text-gray-400 leading-relaxed">{f.d}</p>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* How it works */}
      <section id="how-it-works" className="py-24">
        <div className="max-w-6xl mx-auto px-6">
          <div className="text-center mb-16">
            <h2 className="text-sm font-bold tracking-[0.3em] text-blue-500 mb-2 uppercase">Workflow</h2>
            <h3 className="text-3xl md:text-4xl font-bold text-white">わずか3ステップで完了</h3>
          </div>

          <div className="grid md:grid-cols-3 gap-8">
            {[
              { s: "Step 1", t: "Add Nodes", d: "左側のToolboxから、読み込み(Source)や結合(Join)などのノードをドラッグ＆ドロップで配置します。" },
              { s: "Step 2", t: "Connect Flow", d: "ノード同士の端子をマウスで繋ぎます。データが左から右へと水のように流れて処理されます。" },
              { s: "Step 3", t: "Preview & Export", d: "画面下部に結果がリアルタイム表示されます。グラフ化や、CSV・Excelへのエクスポートが可能です。" }
            ].map((step, idx) => (
              <div key={idx} className="relative">
                {idx !== 2 && <div className="hidden md:block absolute top-8 left-1/2 w-full h-[1px] bg-gradient-to-r from-blue-500/50 to-transparent z-0"></div>}
                <div className="bg-[#182032] border border-gray-800 p-8 rounded-2xl relative z-10 h-full">
                  <div className="text-blue-500 font-bold tracking-widest text-xs mb-2">{step.s}</div>
                  <h4 className="text-xl font-bold text-white mb-4 uppercase">{step.t}</h4>
                  <p className="text-sm text-gray-400 leading-relaxed">{step.d}</p>
                </div>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* CTA / Download */}
      <section id="download" className="py-24 bg-gradient-to-b from-[#0B0F19] to-[#080B13] border-t border-gray-900">
        <div className="max-w-4xl mx-auto px-6 text-center">
          <div className="w-20 h-20 bg-blue-600/20 rounded-full flex items-center justify-center text-blue-400 mx-auto mb-6 border border-blue-500/30">
            <span className="text-4xl">{Icons.Diamond}</span>
          </div>
          <h2 className="text-3xl md:text-5xl font-bold text-white mb-6">さあ、データを整えよう</h2>
          <p className="text-gray-400 mb-10 text-lg">面倒なExcel作業やSQLの記述から解放されます。<br/>完全無料で全機能をご利用いただけます。</p>
          
          <div className="flex flex-col sm:flex-row justify-center gap-4">
            <a href="https://www.vector.co.jp/soft/..." target="_blank" rel="noreferrer" className="bg-blue-600 hover:bg-blue-500 text-white text-sm font-bold px-8 py-4 rounded-xl tracking-widest transition-all shadow-[0_0_20px_rgba(37,99,235,0.4)] flex items-center justify-center gap-2">
              {Icons.Download} Vectorからダウンロード
            </a>
            <a href="https://github.com/your-username/visual-data-prep/releases" target="_blank" rel="noreferrer" className="bg-gray-800 hover:bg-gray-700 text-white text-sm font-bold px-8 py-4 rounded-xl tracking-widest transition-all flex items-center justify-center gap-2 border border-gray-700">
              {Icons.Github} GitHub Releases
            </a>
          </div>
          <p className="text-xs text-gray-600 mt-6 tracking-wider">対応OS: Windows 10 / 11 (64bit)</p>
        </div>
      </section>

      {/* Footer */}
      <footer className="border-t border-gray-900 bg-[#080B13] py-8 text-center">
        <p className="text-xs font-bold tracking-widest text-gray-600 uppercase">
          &copy; {new Date().getFullYear()} Visual Data Prep. All rights reserved.
        </p>
      </footer>
    </div>
  );
}
/* ============================================
   app.js — ActivoRegentis Presentation Logic
   Navigation · Canvas · Touch · Export
   ============================================ */

document.addEventListener('DOMContentLoaded', () => {

    /* ── Elements ───────────────────────────── */
    const slides          = document.querySelectorAll('.slide');
    const prevBtn         = document.getElementById('prevBtn');
    const nextBtn         = document.getElementById('nextBtn');
    const progressBar     = document.getElementById('progressBar');
    const currentSlideNum = document.getElementById('currentSlideNum');
    const totalSlidesNum  = document.getElementById('totalSlidesNum');
    const btnFullscreen   = document.getElementById('btnFullscreen');
    const btnPDF          = document.getElementById('btnPDF');
    const btnPPTX         = document.getElementById('btnPPTX');

    let currentIndex = 0;
    const total = slides.length;
    const isMobile = () => window.innerWidth <= 768;

    if (totalSlidesNum) totalSlidesNum.textContent = total;

    /* ── Slide Navigation ───────────────────── */
    function goTo(index) {
        slides.forEach((slide, i) => {
            slide.classList.remove('active', 'prev');
            if (i === index)      slide.classList.add('active');
            else if (i < index)   slide.classList.add('prev');
        });

        currentIndex = index;

        if (currentSlideNum) currentSlideNum.textContent = index + 1;

        const pct = (index / (total - 1)) * 100;
        if (progressBar) progressBar.style.width = `${pct}%`;

        if (prevBtn) prevBtn.disabled = index === 0;
        if (nextBtn) nextBtn.disabled = index === total - 1;

        // On mobile: scroll the page back to top of content
        if (isMobile()) {
            const deckEl = document.querySelector('.deck-container');
            if (deckEl) {
                window.scrollTo({ top: 0, behavior: 'smooth' });
            }
        }
    }

    function next() { if (currentIndex < total - 1) goTo(currentIndex + 1); }
    function prev() { if (currentIndex > 0)          goTo(currentIndex - 1); }

    if (nextBtn) nextBtn.addEventListener('click', next);
    if (prevBtn) prevBtn.addEventListener('click', prev);

    /* ── Keyboard Navigation ────────────────── */
    document.addEventListener('keydown', (e) => {
        if (['ArrowRight', ' ', 'PageDown'].includes(e.key)) { e.preventDefault(); next(); }
        if (['ArrowLeft', 'Backspace', 'PageUp'].includes(e.key)) { e.preventDefault(); prev(); }
    });

    /* ── Touch / Swipe Navigation ───────────── */
    let touchStartX = 0;
    let touchStartY = 0;
    let touchStartedOnScrollable = false;

    document.addEventListener('touchstart', (e) => {
        touchStartX = e.changedTouches[0].clientX;
        touchStartY = e.changedTouches[0].clientY;

        // Check if the touch started inside a scrollable element
        let el = e.target;
        touchStartedOnScrollable = false;
        while (el && el !== document.body) {
            if (el.scrollHeight > el.clientHeight + 4) {
                touchStartedOnScrollable = true;
                break;
            }
            el = el.parentElement;
        }
    }, { passive: true });

    document.addEventListener('touchend', (e) => {
        if (touchStartedOnScrollable) return;

        const dx = e.changedTouches[0].clientX - touchStartX;
        const dy = e.changedTouches[0].clientY - touchStartY;

        // Only trigger slide change if horizontal swipe is dominant
        if (Math.abs(dx) > Math.abs(dy) && Math.abs(dx) > 55) {
            if (dx < 0) next();
            else         prev();
        }
    }, { passive: true });

    /* ── Fullscreen ─────────────────────────── */
    function updateFullscreenBtn(isActive) {
        if (!btnFullscreen) return;
        const label = btnFullscreen.querySelector('.btn-label');
        const icon  = btnFullscreen.querySelector('.btn-icon');
        if (isActive) {
            if (label) label.textContent = 'Sair';
            if (icon)  icon.textContent  = '✕';
        } else {
            if (label) label.textContent = 'Apresentar';
            if (icon)  icon.textContent  = '⛶';
        }
    }

    if (btnFullscreen) {
        btnFullscreen.addEventListener('click', () => {
            if (!document.fullscreenElement) {
                document.documentElement.requestFullscreen()
                    .then(() => updateFullscreenBtn(true))
                    .catch(err => console.warn('Fullscreen:', err.message));
            } else {
                document.exitFullscreen();
            }
        });
    }

    document.addEventListener('fullscreenchange', () => {
        updateFullscreenBtn(!!document.fullscreenElement);
    });

    /* ── PDF Export ─────────────────────────── */
    if (btnPDF) {
        btnPDF.addEventListener('click', () => window.print());
    }

    /* ── PPTX Export ────────────────────────── */
    if (btnPPTX) {
        btnPPTX.addEventListener('click', () => {
            if (typeof PptxGenJS === 'undefined') {
                alert('Biblioteca PPTX não carregada. Verifique o ficheiro pptxgen.bundle.js.');
                return;
            }

            const pptx = new PptxGenJS();
            pptx.layout = 'LAYOUT_16x9';
            pptx.defineSlideMaster({
                title: 'AR_MASTER',
                background: { color: '000C1A' }
            });

            // ── Helper: standard slide with header ──
            function slide(title, category) {
                const s = pptx.addSlide({ masterName: 'AR_MASTER' });
                // Subtle top accent line
                s.addShape(pptx.shapes.RECTANGLE, {
                    x: 0, y: 0, w: '100%', h: 0.04,
                    fill: { color: 'C5A059' }, line: { color: 'C5A059' }
                });
                if (category) {
                    s.addText(category.toUpperCase(), {
                        x: 0.5, y: 0.3, w: 12.3, h: 0.25,
                        fontSize: 9, fontFace: 'Montserrat',
                        color: 'C5A059', bold: true, charSpacing: 3
                    });
                }
                if (title) {
                    s.addText(title, {
                        x: 0.5, y: 0.55, w: 12.3, h: 0.55,
                        fontSize: 28, fontFace: 'Cinzel', color: 'FFFFFF'
                    });
                    s.addShape(pptx.shapes.RECTANGLE, {
                        x: 0.5, y: 1.12, w: 12.3, h: 0.012,
                        fill: { color: '1A2E4A' }, line: { color: '1A2E4A' }
                    });
                }
                return s;
            }

            // ── Helper: card box ──
            function cardBox(s, x, y, w, h) {
                s.addShape(pptx.shapes.RECTANGLE, {
                    x, y, w, h,
                    fill: { color: '001C3D', transparency: 15 },
                    line: { color: 'C5A059', width: 1 }
                });
            }

            /* ──────────────────────────────────────── */
            /* SLIDE 1 — CAPA                          */
            const s1 = pptx.addSlide({ masterName: 'AR_MASTER' });
            s1.addShape(pptx.shapes.RECTANGLE, {
                x: 0, y: 0, w: '100%', h: 0.04,
                fill: { color: 'C5A059' }, line: { color: 'C5A059' }
            });
            s1.addImage({ path: 'logo.png', x: 4.15, y: 0.9, w: 5.0, h: 4.8 });
            s1.addText('Direito  •  Mercado de Capitais  •  Governação Corporativa', {
                x: 1.0, y: 6.0, w: 11.3, h: 0.4,
                fontSize: 12, fontFace: 'Montserrat',
                color: '718096', align: 'center', charSpacing: 1
            });

            /* ──────────────────────────────────────── */
            /* SLIDE 2 — A NOSSA ESSÊNCIA              */
            const s2 = slide('A Nossa Essência', 'Institucional');
            s2.addText('A ActivoRegentis é especializada na interseção de três pilares fundamentais:', {
                x: 0.5, y: 1.3, w: 12.3, h: 0.45,
                fontSize: 15, fontFace: 'Montserrat', color: 'FFFFFF', align: 'center'
            });
            const s2cards = [
                { title: '⚖  Direito', body: 'O enquadramento legal que assegura a solidez e a legitimidade das operações.' },
                { title: '📈  Mercado de Capitais', body: 'O dinamismo financeiro, investimento e eficiência na alocação de recursos.' },
                { title: '🏛  Governação Corporativa', body: 'Os princípios éticos, a transparência e as melhores práticas de liderança.' },
            ];
            s2cards.forEach((c, i) => {
                const cx = 0.5 + i * 4.3;
                cardBox(s2, cx, 1.9, 4.0, 3.0);
                s2.addText(c.title, { x: cx + 0.15, y: 2.05, w: 3.7, h: 0.5, fontSize: 17, fontFace: 'Cinzel', color: 'C5A059', align: 'center', bold: true });
                s2.addText(c.body, { x: cx + 0.15, y: 2.65, w: 3.7, h: 1.9, fontSize: 13, fontFace: 'Montserrat', color: 'A0AEC0', align: 'center', lineSpacing: 20 });
            });
            s2.addText('Produzimos conteúdos que ajudam profissionais e instituições a tomar decisões mais informadas, estratégicas e sustentáveis.', {
                x: 0.5, y: 5.1, w: 12.3, h: 0.5, fontSize: 13, fontFace: 'Montserrat', color: 'FFFFFF', align: 'center', italic: true
            });

            /* ──────────────────────────────────────── */
            /* SLIDE 3 — O SIGNIFICADO DO NOME         */
            const s3 = slide('O Significado do Nome', 'Identidade');
            cardBox(s3, 0.5, 1.3, 5.5, 3.8);
            s3.addText('ACTIVO', { x: 0.7, y: 1.45, w: 5.1, h: 0.45, fontSize: 22, fontFace: 'Cinzel', color: 'C5A059', bold: true });
            s3.addText('• Ativos e Investimentos\n• Capital produtivo\n• Participação ativa no mercado\n• Dinamismo dos mercados financeiros', {
                x: 0.7, y: 2.0, w: 5.1, h: 2.8, fontSize: 14, fontFace: 'Montserrat', color: 'A0AEC0', lineSpacing: 28
            });
            s3.addText('&', { x: 6.3, y: 2.7, w: 0.7, h: 0.7, fontSize: 32, fontFace: 'Cinzel', color: 'C5A059', align: 'center' });
            cardBox(s3, 7.3, 1.3, 5.5, 3.8);
            s3.addText('REGENTIS', { x: 7.5, y: 1.45, w: 5.1, h: 0.45, fontSize: 22, fontFace: 'Cinzel', color: 'C5A059', bold: true });
            s3.addText('Do latim "Regens"\n\n• Orientar com rigor\n• Estruturar o futuro\n• Dar direção clara\n• Liderar com princípios', {
                x: 7.5, y: 1.98, w: 5.1, h: 2.8, fontSize: 14, fontFace: 'Montserrat', color: 'A0AEC0', lineSpacing: 24
            });
            s3.addText('"ActivoRegentis representa a união perfeita entre a realidade dos mercados e a orientação estratégica baseada no conhecimento."', {
                x: 0.5, y: 5.3, w: 12.3, h: 0.55, fontSize: 13, fontFace: 'Montserrat', color: 'FFFFFF', align: 'center', italic: true
            });

            /* ──────────────────────────────────────── */
            /* SLIDE 4 — A NOSSA MISSÃO                */
            const s4 = slide('A Nossa Missão', 'Propósito');
            cardBox(s4, 0.8, 1.4, 11.7, 4.2);
            s4.addText('"Promover uma cultura de rigor, integridade e excelência na interseção entre o direito, o mercado de capitais e a governação corporativa, através da produção e partilha de conhecimento técnico que gera valor real para profissionais, instituições e para o sistema financeiro."', {
                x: 1.1, y: 1.7, w: 11.1, h: 3.6, fontSize: 20, fontFace: 'Cinzel',
                color: 'FFFFFF', align: 'center', italic: true, lineSpacing: 38
            });

            /* ──────────────────────────────────────── */
            /* SLIDE 5 — A NOSSA VISÃO                 */
            const s5 = slide('A Nossa Visão', 'Futuro');
            s5.addText('Ser uma plataforma de referência, contribuindo ativamente para:', {
                x: 0.5, y: 1.3, w: 12.3, h: 0.45, fontSize: 15, fontFace: 'Montserrat', color: 'FFFFFF', align: 'center'
            });
            const s5items = [
                { t: 'Melhor qualidade das decisões', b: 'Fornecer informação estruturada e análises profundas para mitigar riscos corporativos.' },
                { t: 'Evolução regulatória', b: 'Contribuir para o debate público e técnico sobre novas normas jurídicas e financeiras.' },
                { t: 'Melhores práticas institucionais', b: 'Elevar o nível ético e operacional da governação nos conselhos de administração.' },
                { t: 'Fortalecimento do sistema financeiro', b: 'Fomentar um ecossistema mais transparente, eficiente e atrativo para os investidores.' },
            ];
            s5items.forEach((item, i) => {
                const col = i % 2, row = Math.floor(i / 2);
                const cx = 0.5 + col * 6.5, cy = 1.9 + row * 2.1;
                s5.addShape(pptx.shapes.RECTANGLE, {
                    x: cx, y: cy, w: 0.05, h: 1.8,
                    fill: { color: 'C5A059' }, line: { color: 'C5A059' }
                });
                cardBox(s5, cx + 0.05, cy, 6.0, 1.8);
                s5.addText(item.t, { x: cx + 0.25, y: cy + 0.15, w: 5.7, h: 0.4, fontSize: 14, fontFace: 'Cinzel', color: 'C5A059', bold: true });
                s5.addText(item.b, { x: cx + 0.25, y: cy + 0.6, w: 5.7, h: 1.0, fontSize: 12, fontFace: 'Montserrat', color: 'A0AEC0', lineSpacing: 18 });
            });

            /* ──────────────────────────────────────── */
            /* SLIDE 6 — OS NOSSOS VALORES             */
            const s6 = slide('Os Nossos Valores', 'Princípios');
            const s6vals = [
                { icon: '§', name: 'Rigor',       desc: 'Profundidade técnica, base legal sólida e precisão conceptual em cada publicação.' },
                { icon: '✔', name: 'Integridade', desc: 'Transparência total, independência intelectual e compromisso ético inabalável.' },
                { icon: '★', name: 'Excelência',  desc: 'Melhoria contínua, qualidade profissional de topo e referência técnica.' },
                { icon: '➤', name: 'Impacto',     desc: 'Influenciar decisões estratégicas, melhorar práticas e fortalecer instituições.' },
            ];
            s6vals.forEach((v, i) => {
                const cx = 0.5 + i * 3.2;
                cardBox(s6, cx, 1.4, 2.9, 4.2);
                s6.addText(v.icon, { x: cx, y: 1.55, w: 2.9, h: 0.6, fontSize: 26, color: 'C5A059', align: 'center' });
                s6.addText(v.name.toUpperCase(), { x: cx + 0.1, y: 2.25, w: 2.7, h: 0.45, fontSize: 16, fontFace: 'Cinzel', color: 'C5A059', bold: true, align: 'center' });
                s6.addText(v.desc, { x: cx + 0.1, y: 2.8, w: 2.7, h: 2.5, fontSize: 12, fontFace: 'Montserrat', color: 'A0AEC0', align: 'center', lineSpacing: 20 });
            });

            /* ──────────────────────────────────────── */
            /* SLIDE 7 — ÁREAS DE ACTUAÇÃO             */
            const s7 = slide('Áreas de Actuação', 'Foco');
            const s7areas = ['Direito', 'Mercado de Capitais', 'Governação Corporativa', 'Compliance', 'Regulação Financeira', 'Gestão de Risco', 'Sustentabilidade Institucional', 'Tomada de Decisão Estratégica'];
            s7areas.forEach((area, i) => {
                const col = i % 2, row = Math.floor(i / 2);
                const cx = 0.5 + col * 6.5, cy = 1.4 + row * 1.2;
                s7.addShape(pptx.shapes.RECTANGLE, {
                    x: cx, y: cy + 0.12, w: 0.06, h: 0.7,
                    fill: { color: 'C5A059' }, line: { color: 'C5A059' }
                });
                s7.addShape(pptx.shapes.RECTANGLE, {
                    x: cx + 0.06, y: cy, w: 5.9, h: 1.0,
                    fill: { color: '001C3D', transparency: 10 }, line: { color: '1A2E4A' }
                });
                s7.addText(area, {
                    x: cx + 0.22, y: cy + 0.25, w: 5.6, h: 0.5,
                    fontSize: 16, fontFace: 'Cinzel', color: 'FFFFFF', bold: false
                });
            });

            /* ──────────────────────────────────────── */
            /* SLIDE 8 — O QUE PRODUZIMOS              */
            const s8 = slide('O Que Produzimos', 'Conteúdos');
            const s8prods = ['Artigos Técnicos', 'Análises Jurídicas', 'Insights de Mercado', 'Conteúdo Educativo', 'Opiniões Especializadas', 'Comentários Regulatórios', 'Tendências de Governação', 'Estudos e Reflexões'];
            s8prods.forEach((prod, i) => {
                const col = i % 2, row = Math.floor(i / 2);
                const cx = 0.5 + col * 6.5, cy = 1.4 + row * 1.2;
                s8.addShape(pptx.shapes.RECTANGLE, {
                    x: cx, y: cy + 0.12, w: 0.06, h: 0.7,
                    fill: { color: 'C5A059' }, line: { color: 'C5A059' }
                });
                s8.addShape(pptx.shapes.RECTANGLE, {
                    x: cx + 0.06, y: cy, w: 5.9, h: 1.0,
                    fill: { color: '001C3D', transparency: 10 }, line: { color: '1A2E4A' }
                });
                s8.addText(prod, {
                    x: cx + 0.22, y: cy + 0.25, w: 5.6, h: 0.5,
                    fontSize: 16, fontFace: 'Cinzel', color: 'FFFFFF'
                });
            });

            /* ──────────────────────────────────────── */
            /* SLIDE 9 — PRESENÇA DIGITAL              */
            const s9 = slide('Presença Digital', 'Canais');
            s9.addText('A ActivoRegentis posiciona-se como uma plataforma digital de referência no ecossistema lusófono.', {
                x: 0.5, y: 1.35, w: 12.3, h: 0.45, fontSize: 15, fontFace: 'Montserrat', color: 'FFFFFF', align: 'center'
            });
            const s9channels = [{ name: 'Facebook', x: 1.3 }, { name: 'Instagram', x: 5.0 }, { name: 'LinkedIn', x: 8.7 }];
            s9channels.forEach(ch => {
                cardBox(s9, ch.x, 2.1, 3.2, 2.5);
                s9.addText(ch.name, {
                    x: ch.x + 0.1, y: 3.05, w: 3.0, h: 0.55,
                    fontSize: 20, fontFace: 'Cinzel', color: 'C5A059', align: 'center', bold: true
                });
            });
            s9.addText('Website Oficial:  www.activoregentis.com', {
                x: 0.5, y: 4.9, w: 12.3, h: 0.55, fontSize: 18, fontFace: 'Montserrat', color: 'FFFFFF', align: 'center', bold: true
            });

            /* ──────────────────────────────────────── */
            /* SLIDE 10 — PÚBLICO-ALVO                 */
            const s10 = slide('Público-Alvo', 'Destinatários');
            s10.addText('Levamos conhecimento especializado para moldar o futuro profissional e institucional:', {
                x: 0.5, y: 1.3, w: 12.3, h: 0.45, fontSize: 15, fontFace: 'Montserrat', color: 'A0AEC0', align: 'center'
            });
            const s10targets = ['Profissionais do Direito', 'Administradores', 'Conselheiros de Administração', 'Investidores', 'Reguladores', 'Instituições Financeiras', 'Empresas', 'Académicos', 'Estudantes', 'Mercado de Capitais'];
            s10targets.forEach((t, i) => {
                const col = i % 5, row = Math.floor(i / 5);
                const cx = 0.3 + col * 2.6, cy = 1.95 + row * 1.2;
                cardBox(s10, cx, cy, 2.4, 1.0);
                s10.addText(t, { x: cx + 0.1, y: cy + 0.22, w: 2.2, h: 0.6, fontSize: 11.5, fontFace: 'Montserrat', color: 'FFFFFF', align: 'center', bold: false });
            });

            /* ──────────────────────────────────────── */
            /* SLIDE 11 — O NOSSO DIFERENCIAL          */
            const s11 = slide('O Nosso Diferencial', 'Valor');
            cardBox(s11, 0.8, 1.4, 11.7, 4.3);
            s11.addText('A ActivoRegentis não comunica apenas informação.', {
                x: 1.0, y: 1.6, w: 11.3, h: 0.45, fontSize: 17, fontFace: 'Montserrat', color: 'A0AEC0', align: 'center'
            });
            s11.addText('Transformamos Conhecimento em Orientação.', { x: 1.0, y: 2.25, w: 11.3, h: 0.65, fontSize: 23, fontFace: 'Cinzel', color: 'C5A059', align: 'center', bold: true });
            s11.addText('Análise em Decisão.', { x: 1.0, y: 3.1, w: 11.3, h: 0.65, fontSize: 23, fontFace: 'Cinzel', color: 'C5A059', align: 'center', bold: true });
            s11.addText('Experiência em Valor.', { x: 1.0, y: 3.9, w: 11.3, h: 0.65, fontSize: 23, fontFace: 'Cinzel', color: 'C5A059', align: 'center', bold: true });
            s11.addText('Conhecimento que Orienta Decisões.', {
                x: 1.0, y: 5.05, w: 11.3, h: 0.45, fontSize: 15, fontFace: 'Montserrat', color: 'FFFFFF', align: 'center', bold: true
            });

            /* ──────────────────────────────────────── */
            /* SLIDE 12 — A ASSINATURA DO PROJECTO     */
            const s12 = slide('A Assinatura do Projecto', 'Visão');
            cardBox(s12, 0.5, 1.4, 6.5, 4.2);
            s12.addText('ActivoRegentis é mais do que uma marca corporativa.', {
                x: 0.7, y: 1.6, w: 6.1, h: 0.5, fontSize: 15, fontFace: 'Cinzel', color: 'FFFFFF', bold: true
            });
            s12.addText('É a materialização de uma visão estratégica robusta, construída ao longo de décadas de experiência de excelência nos mercados e no meio jurídico.\n\nRepresenta o ponto de fusão ideal onde o conhecimento e a responsabilidade criam impacto sustentável no mercado moderno.', {
                x: 0.7, y: 2.2, w: 6.1, h: 3.0, fontSize: 13, fontFace: 'Montserrat', color: 'A0AEC0', lineSpacing: 24
            });
            const s12kws = ['Mercado', 'Direito', 'Governação', 'Conhecimento', 'Responsabilidade', 'Impacto'];
            s12kws.forEach((kw, i) => {
                const col = i % 2, row = Math.floor(i / 2);
                const cx = 7.3 + col * 2.7, cy = 1.4 + row * 1.45;
                s12.addShape(pptx.shapes.RECTANGLE, {
                    x: cx, y: cy, w: 2.4, h: 1.2,
                    fill: { color: '001C3D' }, line: { color: 'C5A059', width: 1, dashType: 'dash' }
                });
                s12.addText(kw, { x: cx, y: cy + 0.35, w: 2.4, h: 0.5, fontSize: 14, fontFace: 'Cinzel', color: 'C5A059', align: 'center' });
            });

            /* ──────────────────────────────────────── */
            /* SLIDE 13 — ENCERRAMENTO                 */
            const s13 = pptx.addSlide({ masterName: 'AR_MASTER' });
            s13.addShape(pptx.shapes.RECTANGLE, {
                x: 0, y: 0, w: '100%', h: 0.04,
                fill: { color: 'C5A059' }, line: { color: 'C5A059' }
            });
            s13.addImage({ path: 'logo.png', x: 4.15, y: 0.7, w: 5.0, h: 4.8 });
            s13.addText('Contribuindo para decisões mais sólidas, instituições mais fortes e mercados mais sustentáveis.', {
                x: 1.0, y: 5.65, w: 11.3, h: 0.5, fontSize: 13, fontFace: 'Montserrat', color: 'A0AEC0', align: 'center'
            });
            s13.addText('www.activoregentis.com  |  LinkedIn  |  Facebook  |  Instagram', {
                x: 1.0, y: 6.2, w: 11.3, h: 0.4, fontSize: 11, fontFace: 'Montserrat', color: 'C5A059', align: 'center'
            });

            /* ──────────────────────────────────────── */
            pptx.writeFile({ fileName: 'Apresentacao_ActivoRegentis.pptx' });
        });
    }

    /* ── Canvas Particle Animation ──────────── */
    const canvas = document.getElementById('bgCanvas');
    if (canvas) {
        // Disable on mobile to save battery
        if (isMobile()) {
            canvas.style.display = 'none';
        } else {
            const ctx = canvas.getContext('2d');
            let particles = [];
            const COUNT = 40;

            function resizeCanvas() {
                canvas.width  = window.innerWidth;
                canvas.height = window.innerHeight;
            }

            window.addEventListener('resize', () => {
                resizeCanvas();
                if (isMobile()) canvas.style.display = 'none';
                else            canvas.style.display = '';
            });

            resizeCanvas();

            class Particle {
                constructor() { this.reset(); }
                reset() {
                    this.x  = Math.random() * canvas.width;
                    this.y  = Math.random() * canvas.height;
                    this.vx = (Math.random() - 0.5) * 0.35;
                    this.vy = (Math.random() - 0.5) * 0.35;
                    this.r  = Math.random() * 1.4 + 0.8;
                }
                draw() {
                    ctx.beginPath();
                    ctx.arc(this.x, this.y, this.r, 0, Math.PI * 2);
                    ctx.fillStyle = 'rgba(197, 160, 89, 0.35)';
                    ctx.fill();
                }
                update() {
                    this.x += this.vx;
                    this.y += this.vy;
                    if (this.x < 0 || this.x > canvas.width)  this.vx *= -1;
                    if (this.y < 0 || this.y > canvas.height) this.vy *= -1;
                }
            }

            for (let i = 0; i < COUNT; i++) particles.push(new Particle());

            function animate() {
                ctx.clearRect(0, 0, canvas.width, canvas.height);
                for (let i = 0; i < particles.length; i++) {
                    particles[i].update();
                    particles[i].draw();
                    for (let j = i + 1; j < particles.length; j++) {
                        const dx   = particles[i].x - particles[j].x;
                        const dy   = particles[i].y - particles[j].y;
                        const dist = Math.sqrt(dx * dx + dy * dy);
                        if (dist < 140) {
                            ctx.beginPath();
                            ctx.moveTo(particles[i].x, particles[i].y);
                            ctx.lineTo(particles[j].x, particles[j].y);
                            const alpha = (1 - dist / 140) * 0.07;
                            ctx.strokeStyle = `rgba(197, 160, 89, ${alpha})`;
                            ctx.lineWidth = 0.5;
                            ctx.stroke();
                        }
                    }
                }
                requestAnimationFrame(animate);
            }
            animate();
        }
    }

    /* ── Init ───────────────────────────────── */
    goTo(0);
});

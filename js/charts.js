// Futuristic Glassmorphic Family Finance Tracker - SVG Charts Library (Dynamic Currencies)

window.Charts = {
  /**
   * Renders a glowing glassmorphic donut chart
   * @param {string} containerId 
   * @param {Array<{value: number, color: string, label: string}>} segments 
   * @param {string} currencySymbol
   */
  renderDonutChart(containerId, segments, currencySymbol = '₹') {
    const container = document.getElementById(containerId);
    if (!container) return;

    const total = segments.reduce((acc, s) => acc + s.value, 0);
    if (total === 0) {
      container.innerHTML = `<div class="no-data">No Expense Data Available</div>`;
      return;
    }

    // Sort segments so largest is the outermost ring
    const activeSegments = segments.filter(s => s.value > 0).sort((a, b) => b.value - a.value);

    const cx = 50;
    const cy = 50;
    const strokeWidth = 5.2;

    let svgContent = `
      <svg viewBox="0 0 100 100" class="donut-svg">
        <defs>
          <linearGradient id="glow-grad" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="var(--apple-blue)" stop-opacity="0.15"></stop>
            <stop offset="100%" stop-color="var(--apple-purple)" stop-opacity="0.05"></stop>
          </linearGradient>
        </defs>
    `;

    activeSegments.forEach((seg, idx) => {
      // Outermost ring has radius 40, next has 32.5, 25, 17.5
      const radius = 40 - (idx * 7.5);
      if (radius < 10) return; // avoid too tiny rings
      
      const circumference = 2 * Math.PI * radius;
      const percentage = seg.value / total;
      const strokeLength = percentage * circumference;
      const strokeOffset = circumference - strokeLength;

      svgContent += `
        <!-- Background Track Ring -->
        <circle cx="${cx}" cy="${cy}" r="${radius}" 
                stroke="var(--apple-border)" 
                stroke-width="${strokeWidth}" 
                stroke-opacity="0.25"
                fill="transparent" />
        <!-- Active Progress Ring -->
        <circle cx="${cx}" cy="${cy}" r="${radius}" 
                stroke="${seg.color}" 
                stroke-width="${strokeWidth}" 
                stroke-dasharray="${circumference}" 
                stroke-dashoffset="${strokeOffset}" 
                stroke-linecap="round"
                fill="transparent" 
                transform="rotate(-90 ${cx} ${cy})"
                class="donut-segment"
                style="--seg-color: ${seg.color}; transition: stroke-dashoffset 0.8s ease-in-out;"
                data-label="${seg.label}"
                data-value="${seg.value}"
                data-percent="${(percentage * 100).toFixed(0)}%">
          <title>${seg.label}: ${currencySymbol}${seg.value.toFixed(2)} (${(percentage * 100).toFixed(1)}%)</title>
        </circle>
      `;
    });

    // Add center summary text
    svgContent += `
        <circle cx="${cx}" cy="${cy}" r="11" fill="var(--apple-bg)" stroke="var(--apple-border)" stroke-width="1" />
        <text x="${cx}" y="${cy - 2.5}" text-anchor="middle" class="donut-center-title" style="font-size: 3.5px; font-weight: 700; fill: var(--apple-text-secondary); letter-spacing: 0.3px;">EXPENSES</text>
        <text x="${cx}" y="${cy + 3.5}" text-anchor="middle" class="donut-center-value" style="font-size: 5px; font-weight: 800; fill: var(--apple-text);">${currencySymbol}${total.toLocaleString(undefined, {maximumFractionDigits: 0})}</text>
      </svg>
    `;

    container.innerHTML = svgContent;
  },

  /**
   * Renders a glowing cyber line chart for monthly/daily expense trends
   * @param {string} containerId 
   * @param {Array<{label: string, value: number}>} points 
   * @param {string} lineColor 
   * @param {string} currencySymbol
   */
  renderLineChart(containerId, points, lineColor = '#00f0ff', currencySymbol = '₹') {
    const container = document.getElementById(containerId);
    if (!container) return;

    if (!points || points.length === 0) {
      container.innerHTML = `<div class="no-data">No Trend Data Available</div>`;
      return;
    }

    const width = 500;
    const height = 220;
    const paddingX = 45; // slightly wider padding for larger currency strings (e.g. ₹10,000)
    const paddingY = 30;

    const maxVal = Math.max(...points.map(p => p.value), 100) * 1.1; // Add 10% headroom
    const minVal = 0;

    // Generate SVG coordinates
    const svgPoints = points.map((p, i) => {
      const x = paddingX + (i * (width - 2 * paddingX) / (points.length - 1 || 1));
      const y = height - paddingY - ((p.value - minVal) * (height - 2 * paddingY) / (maxVal - minVal));
      return { x, y, label: p.label, value: p.value };
    });

    // Line path d attribute
    let pathD = '';
    if (svgPoints.length > 0) {
      pathD = `M ${svgPoints[0].x} ${svgPoints[0].y}`;
      for (let i = 1; i < svgPoints.length; i++) {
        // Use bezier curves for smooth styling
        const p0 = svgPoints[i - 1];
        const p1 = svgPoints[i];
        const cpX1 = p0.x + (p1.x - p0.x) / 2;
        const cpY1 = p0.y;
        const cpX2 = p0.x + (p1.x - p0.x) / 2;
        const cpY2 = p1.y;
        pathD += ` C ${cpX1} ${cpY1}, ${cpX2} ${cpY2}, ${p1.x} ${p1.y}`;
      }
    }

    // Area path (under the line, closed at the bottom)
    let areaD = '';
    if (svgPoints.length > 0) {
      areaD = `${pathD} L ${svgPoints[svgPoints.length - 1].x} ${height - paddingY} L ${svgPoints[0].x} ${height - paddingY} Z`;
    }

    // Generate Y-axis gridlines & labels
    let gridLines = '';
    const yTicks = 4;
    for (let i = 0; i <= yTicks; i++) {
      const val = minVal + (i * (maxVal - minVal) / yTicks);
      const y = height - paddingY - (i * (height - 2 * paddingY) / yTicks);
      
      // format values to K/M format if huge (e.g. ₹50K instead of ₹50000)
      let displayVal = Math.round(val);
      if (displayVal >= 1000000) {
        displayVal = (displayVal / 1000000).toFixed(1) + 'M';
      } else if (displayVal >= 1000) {
        displayVal = (displayVal / 1000).toFixed(0) + 'K';
      }
      
      gridLines += `
        <line x1="${paddingX}" y1="${y}" x2="${width - paddingX}" y2="${y}" stroke="rgba(255, 255, 255, 0.05)" stroke-dasharray="4 4" />
        <text x="${paddingX - 8}" y="${y + 4}" text-anchor="end" class="chart-axis-text">${currencySymbol}${displayVal}</text>
      `;
    }

    // Generate X-axis labels
    let xLabels = '';
    svgPoints.forEach((pt, i) => {
      const showLabel = points.length <= 7 || i % Math.ceil(points.length / 6) === 0 || i === points.length - 1;
      if (showLabel) {
        xLabels += `
          <text x="${pt.x}" y="${height - paddingY + 18}" text-anchor="middle" class="chart-axis-text">${pt.label}</text>
        `;
      }
    });

    // Glowing points
    let dataPointsSvg = '';
    svgPoints.forEach((pt) => {
      dataPointsSvg += `
        <g class="chart-point-group">
          <circle cx="${pt.x}" cy="${pt.y}" r="8" fill="${lineColor}" opacity="0" class="chart-point-hover" />
          <circle cx="${pt.x}" cy="${pt.y}" r="4" fill="#ffffff" stroke="${lineColor}" stroke-width="2" class="chart-point" />
          <title>${pt.label}: ${currencySymbol}${pt.value.toLocaleString()}</title>
        </g>
      `;
    });

    let svgContent = `
      <svg viewBox="0 0 ${width} ${height}" class="line-chart-svg">
        <defs>
          <!-- Line Gradient -->
          <linearGradient id="line-grad" x1="0%" y1="0%" x2="100%" y2="0%">
            <stop offset="0%" stop-color="${lineColor}" />
            <stop offset="100%" stop-color="#ff007f" />
          </linearGradient>
          <!-- Area Fill Gradient -->
          <linearGradient id="area-grad" x1="0%" y1="0%" x2="0%" y2="100%">
            <stop offset="0%" stop-color="${lineColor}" stop-opacity="0.25" />
            <stop offset="100%" stop-color="${lineColor}" stop-opacity="0.0" />
          </linearGradient>
          <!-- Glow Filter -->
          <filter id="chart-glow" x="-20%" y="-20%" width="140%" height="140%">
            <feGaussianBlur stdDeviation="4" result="blur" />
            <feMerge>
              <feMergeNode in="blur" />
              <feMergeNode in="SourceGraphic" />
            </feMerge>
          </filter>
        </defs>

        <!-- Grid lines & Y axis -->
        ${gridLines}

        <!-- Area Fill -->
        ${areaD ? `<path d="${areaD}" fill="url(#area-grad)" />` : ''}

        <!-- Trend Line -->
        ${pathD ? `<path d="${pathD}" fill="none" stroke="url(#line-grad)" stroke-width="3.5" filter="url(#chart-glow)" stroke-linecap="round" />` : ''}

        <!-- Interactive Points -->
        ${dataPointsSvg}

        <!-- X axis labels -->
        ${xLabels}
      </svg>
    `;

    container.innerHTML = svgContent;
  },

  /**
   * Generates an SVG radial progress ring for budgets and saving goals
   * @param {number} current 
   * @param {number} target 
   * @param {string} color 
   * @returns {string} SVG HTML string
   */
  getRadialProgressSVG(current, target, color = '#00f0ff') {
    const percent = Math.min((current / target) * 100, 100);
    const radius = 15;
    const circumference = 2 * Math.PI * radius; // ~94.25
    const offset = circumference - (percent / 100) * circumference;

    return `
      <svg width="36" height="36" viewBox="0 0 36 36" class="radial-progress-svg">
        <!-- Background track -->
        <circle cx="18" cy="18" r="${radius}" 
                stroke="rgba(255, 255, 255, 0.05)" 
                stroke-width="3" 
                fill="transparent" />
        <!-- Colored indicator -->
        <circle cx="18" cy="18" r="${radius}" 
                stroke="${color}" 
                stroke-width="3" 
                stroke-dasharray="${circumference}" 
                stroke-dashoffset="${offset}" 
                stroke-linecap="round"
                fill="transparent" 
                transform="rotate(-90 18 18)"
                style="filter: drop-shadow(0 0 3px ${color});" />
      </svg>
    `;
  }
};

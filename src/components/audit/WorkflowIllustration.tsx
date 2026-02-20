export function WorkflowIllustration() {
  return (
    <div className="flex flex-col items-center py-6">
      <svg width="600" height="140" viewBox="0 0 600 140" fill="none" xmlns="http://www.w3.org/2000/svg">
        {/* Step 1 - Upload Template */}
        <g>
          <rect x="10" y="20" width="140" height="90" rx="12" fill="hsl(222, 100%, 28%)" fillOpacity="0.08" stroke="hsl(222, 100%, 28%)" strokeWidth="2" strokeDasharray="6 3" />
          <rect x="55" y="35" width="50" height="40" rx="4" fill="hsl(222, 100%, 28%)" fillOpacity="0.15" />
          <path d="M70 60 L80 50 L90 60" stroke="hsl(222, 100%, 28%)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
          <line x1="80" y1="50" x2="80" y2="68" stroke="hsl(222, 100%, 28%)" strokeWidth="2" strokeLinecap="round" />
          <text x="80" y="100" textAnchor="middle" fill="hsl(222, 100%, 28%)" fontSize="11" fontWeight="600" fontFamily="Arial">Upload Template</text>
        </g>

        {/* Arrow 1 */}
        <g>
          <line x1="160" y1="65" x2="210" y2="65" stroke="hsl(222, 100%, 28%)" strokeWidth="2" strokeDasharray="4 4" />
          <polygon points="210,60 220,65 210,70" fill="hsl(222, 100%, 28%)" />
        </g>

        {/* Step 2 - Upload Excel */}
        <g>
          <rect x="230" y="20" width="140" height="90" rx="12" fill="hsl(142, 76%, 36%)" fillOpacity="0.08" stroke="hsl(142, 76%, 36%)" strokeWidth="2" strokeDasharray="6 3" />
          <rect x="275" y="35" width="50" height="40" rx="4" fill="hsl(142, 76%, 36%)" fillOpacity="0.15" />
          <rect x="282" y="42" width="36" height="4" rx="1" fill="hsl(142, 76%, 36%)" fillOpacity="0.6" />
          <rect x="282" y="50" width="24" height="4" rx="1" fill="hsl(142, 76%, 36%)" fillOpacity="0.4" />
          <rect x="282" y="58" width="30" height="4" rx="1" fill="hsl(142, 76%, 36%)" fillOpacity="0.5" />
          <text x="300" y="100" textAnchor="middle" fill="hsl(142, 76%, 36%)" fontSize="11" fontWeight="600" fontFamily="Arial">Upload Excel</text>
        </g>

        {/* Arrow 2 */}
        <g>
          <line x1="380" y1="65" x2="430" y2="65" stroke="hsl(38, 92%, 50%)" strokeWidth="2" strokeDasharray="4 4" />
          <polygon points="430,60 440,65 430,70" fill="hsl(38, 92%, 50%)" />
        </g>

        {/* Step 3 - Generated PPT */}
        <g>
          <rect x="450" y="20" width="140" height="90" rx="12" fill="hsl(38, 92%, 50%)" fillOpacity="0.08" stroke="hsl(38, 92%, 50%)" strokeWidth="2" />
          <rect x="495" y="32" width="50" height="40" rx="4" fill="hsl(38, 92%, 50%)" fillOpacity="0.2" />
          <path d="M510 47 L520 42 L530 47 L530 62 L510 62 Z" fill="hsl(38, 92%, 50%)" fillOpacity="0.35" />
          <circle cx="515" cy="56" r="3" fill="hsl(38, 92%, 50%)" fillOpacity="0.5" />
          <text x="520" y="90" textAnchor="middle" fill="hsl(38, 92%, 50%)" fontSize="11" fontWeight="600" fontFamily="Arial">Auto PPT ✨</text>
          <text x="520" y="108" textAnchor="middle" fill="hsl(38, 92%, 50%)" fontSize="10" fontFamily="Arial">Voilà!</text>
        </g>

        {/* Character - small celebrating person */}
        <g transform="translate(565, 25)">
          <circle cx="15" cy="8" r="7" fill="hsl(222, 100%, 28%)" fillOpacity="0.2" stroke="hsl(222, 100%, 28%)" strokeWidth="1.5" />
          <line x1="15" y1="15" x2="15" y2="35" stroke="hsl(222, 100%, 28%)" strokeWidth="2" strokeLinecap="round" />
          <line x1="15" y1="20" x2="5" y2="12" stroke="hsl(222, 100%, 28%)" strokeWidth="2" strokeLinecap="round" />
          <line x1="15" y1="20" x2="25" y2="12" stroke="hsl(222, 100%, 28%)" strokeWidth="2" strokeLinecap="round" />
          <line x1="15" y1="35" x2="8" y2="48" stroke="hsl(222, 100%, 28%)" strokeWidth="2" strokeLinecap="round" />
          <line x1="15" y1="35" x2="22" y2="48" stroke="hsl(222, 100%, 28%)" strokeWidth="2" strokeLinecap="round" />
          {/* sparkles */}
          <text x="0" y="5" fontSize="8">✨</text>
          <text x="24" y="5" fontSize="8">🎉</text>
        </g>
      </svg>
    </div>
  );
}

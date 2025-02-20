class Header():

    def fte0():
        return r"""
        \specialrule{1.5pt}{0pt}{0pt}
        \multicolumn{1}{|l}{
        \textbf{\shortstack[l]{
        \rule{0pt}{2em} % Top space adjustment
        Department \# - Department Name\\
        \hspace{0.5cm}Fund \# - Fund Name\\
        \hspace{1cm}Appropriation \# - Appropriation Name\\
        \hspace{1.5cm}Cost Center \# - Cost Center Name\\
        \hspace{2cm}Job Code \# - Job Title}}
        \rule[-1em]{0pt}{4em} % Bottom space adjustment
        } &
        \textbf{\shortstack{FY2025 \\ Adopted}} &
        \textbf{\shortstack{FY2026 \\ Adopted}} &
        \textbf{\shortstack{FY2027 \\ Forecast}} &
        \textbf{\shortstack{FY2028 \\ Forecast}} &
        \multicolumn{1}{c|}{
        \textbf{\shortstack{FY2029 \\ Forecast}}
        } \\
        \specialrule{1.5pt}{0pt}{0pt}

\endfirsthead
     
        \multicolumn{6}{c}{\textit{
            \textnormal{\textbf{\shortstack[c]{
                CITY OF DETROIT%
                \\ {BUDGET DEVELOPMENT}%
                \\ {POSITION DETAIL BY DEPARTMENT, FUND, APPROPRIATION, \& COST CENTER}%
                \\ {DEPARTMENT 72 - DETROIT PUBLIC LIBRARY}}
            }
            \rule[-1em]{0pt}{6em} % Bottom space adjustment
        }}} \\        
        \specialrule{1.5pt}{0pt}{0pt}
        \multicolumn{1}{|l}{
        \textbf{\shortstack[l]{
        \rule{0pt}{2em} % Top space adjustment
        Department \# - Department Name\\
        \hspace{0.5cm}Fund \# - Fund Name\\
        \hspace{1cm}Appropriation \# - Appropriation Name\\
        \hspace{1.5cm}Cost Center \# - Cost Center Name\\
        \hspace{2cm}Job Code \# - Job Title}}
        \rule[-1em]{0pt}{4em} % Bottom space adjustment
        } &
        \textbf{\shortstack{FY2025 \\ Adopted}} &
        \textbf{\shortstack{FY2026 \\ Adopted}} &
        \textbf{\shortstack{FY2027 \\ Forecast}} &
        \textbf{\shortstack{FY2028 \\ Forecast}} &
        \multicolumn{1}{c|}{
        \textbf{\shortstack{FY2029 \\ Forecast}}
        } \\
        \specialrule{1.5pt}{0pt}{0pt}

\endhead"""


    def text(main, subheaders,
             fund,
             category,
             approp,
             cc,
             jobName):        
            
        title = r"""
            \specialrule{1.5pt}{0pt}{0pt}
            \multicolumn{1}{|l}{
            \textbf{\shortstack[l]{
                \rule{0pt}{2em} % Top space adjustment
                Department \# - Department Name\\"""
        
        if fund:
            title += r'\hspace{0.5cm}Fund \# - Fund Name\\'

        if category:
            title += r'\hspace{1cm}Summary Category\\'

        if approp:
            title += r'\hspace{1cm}Appropriation \# - Appropriation Name\\'

        if cc:
            title += r'\hspace{1.5cm}Cost Center \# - Cost Center Name\\'

        if jobName:
            title += r'\hspace{2cm}Job Code \# - Job Title'

        title += r"""}}
            \rule[-1em]{0pt}{4em} % Bottom space adjustment
            } &
            \textbf{\shortstack{FY2025 \\ Adopted}} &
            \textbf{\shortstack{FY2026 \\ Adopted}} &
            \textbf{\shortstack{FY2027 \\ Forecast}} &
            \textbf{\shortstack{FY2028 \\ Forecast}} &
            \multicolumn{1}{c|}{
            \textbf{\shortstack{FY2029 \\ Forecast}}
            } \\
            \specialrule{1.5pt}{0pt}{0pt}
            """
        
        firsthead = title + '\n' + r'\endfirsthead'

        top_lines = rf"""
            {main} \\
            {subheaders[0]} \\
            {subheaders[1]} \\
            {subheaders[2]}
        """ + r'}}'
            
        
        head = rf"""
            \multicolumn{{6}}{{c}}{{
                \textit{{
                    \textnormal{{
                        \textbf{{
                            \shortstack[c]{{{top_lines}
                \rule[-1em]{{0pt}}{{6em}} % Bottom space adjustment
            }} }} }} \\        
            {title}
            \endhead
            """
        return firsthead + head

    @staticmethod
    def fte(main, subheaders):
        return Header.text(main,
                         subheaders,
                         fund=True,
                         category=False,
                         approp=True,
                         cc = True,
                         jobName=True
                         )
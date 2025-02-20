class Header():
    """ Create header for first page and every page after """
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
                Department \# - Department Name"""
        
        if fund:
            title += r'\\ \hspace{0.5cm}Fund \# - Fund Name'

        if category:
            title += r'\\ \hspace{1cm}Summary Category'

        if approp:
            title += r'\\ \hspace{1cm}Appropriation \# - Appropriation Name'

        if cc:
            title += r'\\ \hspace{1.5cm}Cost Center \# - Cost Center Name'

        if jobName:
            title += r'\\ \hspace{2cm}Job Code \# - Job Title'

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
        """ header for FTE table of table titles """
        return Header.text(main,
                         subheaders,
                         fund=True,
                         category=False,
                         approp=True,
                         cc = True,
                         jobName=True
                         )
    
    @staticmethod
    def fund_categories(main, subheaders):
        """ header for fund-summary tables """
        return Header.text(main,
                         subheaders,
                         fund=True,
                         category=True,
                         approp=False,
                         cc = False,
                         jobName=False
                         )
    
    @staticmethod
    def summary_categories(main, subheaders):
        """ header for category summary tables """
        return Header.text(main,
                         subheaders,
                         fund=False,
                         category=True,
                         approp=False,
                         cc = False,
                         jobName=False
                         )
    
    @staticmethod
    def fund_approp_cc(main, subheaders):
        """ header for category summary tables """
        return Header.text(main,
                         subheaders,
                         fund=True,
                         category=False,
                         approp=True,
                         cc = True,
                         jobName=False
                         )
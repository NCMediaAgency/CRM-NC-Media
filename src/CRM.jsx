import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { applicaMeta, suggerisciAccount, ACCOUNT_STATI, accountUtilizzabile } from "./lib/meta.js";
import { trovaTabelleKpi, applicaKpi, VOCI_KPI } from "./lib/kpi.js";
import { applicaComandi, statoEsterno, nomeChiave } from "./lib/comandi.js";
import { fusoDelBrowser } from "./lib/spazioApi.js";
import { BENCH_BASE, ORDINI, benchmarkDi, classifica, fotografie, conAndamento, raggruppaCopie } from "./lib/winner.js";
import {
  Wallet, Users, Target, MousePointerClick, TrendingUp, Euro, CheckCircle2,
  AlertTriangle, Clock, BarChart3, Percent, Layers, Briefcase, Gauge,
  LayoutDashboard, Building2, Megaphone, CalendarRange, PieChart, PhoneCall,
  Handshake, Calculator, Filter, NotebookPen, FolderOpen, Sparkles, ListChecks,
  Zap, Upload, DownloadCloud, HelpCircle, Settings, FileText, PenLine, Contact,
  KanbanSquare, Receipt, Code2, BellRing, LogOut, Trophy, Moon, Sun, GraduationCap,
} from "lucide-react";

/* Logo dell'agenzia, incorporato: nessun file esterno da caricare */
const LOGO_AG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAMAAAD04JH5AAAYT2lDQ1BJQ0MgUHJvZmlsZQAAeJyVWXc8le/7v5+zz7HPsffeZO+9994kHOtYccxQIskqUSFKJZmVSqESkYZSRh+SJJJRKVRURn6PUZ/P9/P94/f63q/X/Tzvc93XfY37vu5xPQcArv2+kZFhCEYAwiNiqA6mBvxu7h782HcAAjjAAGiBvC85OlLfzs4KwOX3+x8FAmBpcOMJwHOZDVngfytM/gHRZFiMHYz9/KPJ4TC+DgAqnRxJjQEAowbTheJjIjewF4yZqbCBMI7cwEFbOH0D+23h4k0eJwdDGF8EAEfr60sNAoC+Gabzx5GDYDn0Q3AbMcKfEgGzzsJYhxzs6w8AlzTMIx0evnsDu8FY3O8fcoL+Q6bfH5m+vkF/8JYvmwVnRImODPPd8z8Ox/9fwsNif+sQhSttMNXMYcNneNyGQndbbmBaGM9G+NnYwpgI4x8U/01+GCMIwbFmzlv8CG5ytCE8ZoAVxnL+vkaWMOaGsUlEmI3VNt0vkGJiDmM4QhAJlBhzp+2+mQHRxo7bMk9RdzvY/saBVEP97b6Xfambejf4O2NDnfW35Q8FB5j/lv8tMdjJdctmJCGO4mIDY3oYs0aHOlpu8SCFE4MNbX7zUGMdNuwXhrFGQISpwZZ8pHcg1cRhm58aHv3bX2RmMMXcZhuXxAQ7mW3LuUj23bSfHcbNARH6zr/lBES7Wf32xT/AyHjLd2RvQITztr/IscgYA4ftvl8iw+y2+VGEgDDTDbogjLmj4xy3+6J0YuCA3JKPsomMsXPashPlF+JrYbdlDyoBWAFDYAT4QSxc/cBuEAIoz2abZuFfWy0mwBdQQRAIADLblN89XDdbIuCnI0gEn2AUAKL/9DPYbA0AcTB97Q916ykDAjdb4zZ7hIIpGIcDSxAG/47d7BXxR5sLeAdTKP+lnQzbGgbXjbb/punDFKttSuxvufwMvzkxxhgjjBnGBCOB4kTpoDRRVvBTD64KKDWU+m9r/+ZHT6H70G/RA+gx9MtdlDTqv/zhB9ZgDNZgsu2z3z99RonCUpVRBihtWD4sG8WK4gQyKCVYkz5KF9atDFMNty3f8P7fsv/Dh3+M+jYfXg6PwLPh9fDi/+5JL0mv/EfKxpj+c4S2bPX7M66Gf1r+rd/wHyPtD78t/82JzEQ2IB8i25FdyBZkE+BHtiGbkd3IOxv4TxS924yi39ocNu0JheVQ/kuf77bOjZGMlquTey+3utUWE5AQs7HADHdH7qFSgoJj+PXhnT+A3zyCLCvNryCnIAfAxjmytU19ddg8HyDWnr9p5IMAqM4DgF/+mxb+FYArBHgbtf6bJuINLzMMANVT5Fhq3BYNtfFAAwJ8PjEDDsALhIA47I8CUAGaQA8YAwtgC5yAO/CGRzkYjmcqiAfJIBVkgBxwFJwAJeAMOA+qwSVwDTSBFtAOHoAnoBcMgFdw9EyCj2AeLIEVCIKwEB1EgjggPkgEkoIUIDVIBzKGrCAHyB3ygYKgCCgWSoYOQDlQAVQCnYNqoKvQTagd6oL6oJfQOPQe+gItI5AIWgQzggchitiBUEPoIywRToidiCBEFCIRkY44gihGlCMuIhoR7YgniAHEGOIjYhEJkDRIVqQAUgaphjRE2iI9kIFIKnIfMhtZiCxHXkbeguf5OXIMOYv8icKgSCh+lAwcwWYoZxQZFYXah8pFlaCqUY2oTtRz1DhqHvULTYfmRkuhNdDmaDd0EDoenYEuRFeib6Dvw6tpEr2EwWBYMWIYVXg1umNCMEmYXMxpTD3mLqYPM4FZxGKxHFgprDbWFuuLjcFmYE9iL2LbsP3YSewPHA2OD6eAM8F54CJwabhCXC2uFdePm8at4BnxIngNvC3eH78Hn4evwN/C9+An8SsEJoIYQZvgRAghpBKKCZcJ9wkjhK80NDSCNOo09jQUmv00xTRXaB7RjNP8pCXSStIa0nrRxtIeoa2ivUv7kvYrHR2dKJ0enQddDN0Ruhq6e3SjdD/oSfSy9Ob0/vQp9KX0jfT99J8Z8AwiDPoM3gyJDIUMDQw9DLOMeEZRRkNGX8Z9jKWMNxlfMC4ykZjkmWyZwplymWqZuphmiFiiKNGY6E9MJ54n3iNOkJAkIZIhiUw6QKog3SdNMmOYxZjNmUOYc5gvMT9jnmchsiixuLAksJSy3GEZY0WyirKas4ax5rFeYx1kXWbjYdNnC2DLYrvM1s/2nZ2LXY89gD2bvZ59gH2Zg5/DmCOUI5+jieM1J4pTktOeM56zjPM+5ywXM5cmF5krm+sa1zA3gluS24E7ifs8dzf3Ig8vjylPJM9Jnns8s7ysvHq8IbzHeVt53/OR+HT4KHzH+dr4PvCz8Ovzh/EX83fyzwtwC5gJxAqcE3gmsCIoJugsmCZYL/haiCCkJhQodFyoQ2hemE/YWjhZuE54WAQvoiYSLFIk8lDku6iYqKvoIdEm0RkxdjFzsUSxOrERcTpxXfEo8XLxvyQwEmoSoRKnJXolEZLKksGSpZI9UggpFSmK1GmpPmm0tLp0hHS59AsZWhl9mTiZOplxWVZZK9k02SbZzzuEd3jsyN/xcMcvOWW5MLkKuVfyRHkL+TT5W/JfFCQVyAqlCn8p0imaKKYoNisuKEkpBSiVKQ0pk5StlQ8pdyivqaiqUFUuq7xXFVb1UT2l+kKNWc1OLVftkTpa3UA9Rb1F/aeGikaMxjWNOU0ZzVDNWs0ZLTGtAK0KrQltQW1f7XPaYzr8Oj46Z3XGdAV0fXXLdd/qCen561XqTetL6IfoX9T/bCBnQDW4YfDdUMNwr+FdI6SRqVG20TNjorGzcYnxqImgSZBJncm8qbJpkuldM7SZpVm+2QtzHnOyeY35vIWqxV6LTktaS0fLEsu3VpJWVKtb1ghrC+tj1iM2IjYRNk22wNbc9pjtazsxuyi72/YYezv7UvspB3mHZIeHjiTHXY61jktOBk55Tq+cxZ1jnTtcGFy8XGpcvrsauRa4jrntcNvr9sSd053i3uyB9XDxqPRY9DT2POE56aXsleE1uFNsZ8LOLm9O7zDvO7sYdvnuavBB+7j61Pqs+tr6lvsu+pn7nfKbJxuSi8gf/fX8j/u/D9AOKAiYDtQOLAicCdIOOhb0Plg3uDB4lmJIKaEshJiFnAn5HmobWhW6HuYaVh+OC/cJvxlBjAiN6NzNuzthd1+kVGRG5FiURtSJqHmqJbUyGoreGd0cwwxf2LtjxWMPxo7H6cSVxv2Id4lvSGBKiEjo3iO5J2vPdKJJ4oUkVBI5qSNZIDk1eXyv/t5z+6B9fvs6UoRS0lMm95vur04lpIamPk2TSytI+3bA9cCtdJ70/ekTB00P1mXQZ1AzXhzSPHQmE5VJyXyWpZh1MutXtn/24xy5nMKc1Vxy7uPD8oeLD68fCTzyLE8lr+wo5mjE0cF83fzqAqaCxIKJY9bHGo/zH88+/u3ErhNdhUqFZ4oIRbFFY8VWxc0nhU8ePblaElwyUGpQWn+K+1TWqe+n/U/3l+mVXT7DcybnzPJZytmhc6bnGstFywvPY87HnZ+qcKl4eEHtQk0lZ2VO5VpVRNVYtUN1Z41qTU0td21eHaIutu79Ra+LvZeMLjVflrl8rp61PucKuBJ75cNVn6uD1yyvdTSoNVy+LnL91A3SjexGqHFP43xTcNNYs3tz302Lmx23NG/duC17u6pFoKX0DsudvFZCa3rrelti2+LdyLuz7UHtEx27Ol7dc7v3V6d957P7lvcfPTB5cO+h/sO2R9qPWro0um4+Vnvc9ETlSWO3cveNp8pPbzxTedbYo9rT3Kvee6tPq6+1X7e//bnR8wd/mf/1ZMBmoG/QeXDohdeLsSH/oZmXYS8XhuOGV17tH0GPZL9mfF04yj1a/kbiTf2YytidcaPx7reOb19NkCc+vot+tzqZPkU3VTjNN10zozDT8t7kfe8Hzw+THyM/rsxmfGL6dOqz+Ofrc3pz3fNu85ML1IX1L7lfOb5WfVP61rFotzi6FL608j37B8eP6p9qPx8uuy5Pr8SvYleL1yTWbv2y/DWyHr6+HulL9d28CiDhiggMBOBLFQB07gCQegEgeG7ledsFCV8+EPDbBZKFPiLS4RO1B5WBNsEgMU+wxbgIvBVBggZLM0vbT9dEX8VQyVjP1EzsID1h7mUZYn3DNsP+kWOBc5lrjQfBi+Uj8NMJEAWJQqzC7CJsouxi3OI8EvyS/FKC0sIyorJiO6Tl5OQVFVQUNZR0lY1VzFXN1UzUTTRMNA219LW1dDR0lfRk9UUNeAyZjQhG68ZfTaZMX5p1m7dYVFses0qxDrFxszW2U7YXc+ByZHTCOSNdIFeEG8od78HoyeElvFPGW2KXsA+fL6cfC5nkTwwgBbIGcQULUqRDVENNwlzCKRHJuwsiK6LOUouj82NyY7PisuOPJBTvqU5sTXq1F+yTTtm1/2TqqwOC6bsPth/CZAplKWQb5DjmBh5OPJKfV330bv5wweJxphMyhRZFgcUHTpaV3CztP/Xu9OIZ7FmOc5LlWudtK/wuxFQerCqsrq65Wfu4bvjih0s/63FX2K6KX9NtcL8edSOr8XRTfXPbza5bPbd7W57c6Wi92lZ6N6V9V4fGPeK9qc6b92sfnHqY8yihy++x+RPZbvru2af3n53qiew16CP1TfRfe576l/2AyCBq8P2L7qH6lwXDMa9cRtRec75eHR190z52YTzr7e4J53dak8JwlC1N/zVz/X3Rh5SPYbPkT+TPkXM58zcW5r7qfTu3RPpe/FNq+dlqyi+N9fV/zL8CcgZVgLbEsGBeYxtwufggghGNJC0D7SrdNP0QwxDjG6Z3xE+kr8xLLGusK2xr7L841jiXuL5yz/FM8Y7w9fPfF7gpWCmUIxwmYiUqKYYX+yDeJVEjmS1FkbaUkZGlk53b0Sd3Xb5IIVmRrGSvbKCioCqgRlRbV/+sMaLZpdWoXa6Tqxuv56NvYaBgyGmEMHpv/Mzkimm+WbS5k4WKJZvlitUb63s2tbb5dkn2gQ6OjvpO8s4CLiRXrOuy20f3EY9uzzte9TvPeh/bdcgn2ZfqRyH7+nsEOAXaB9kEW1IsQ8xCNcNkwwUiWHbTRCIiV6N+UH9Gr8Wi44jxQgkae5wSo5MKk1v2TqXQ7OdLlUnTPmCT7ncwPuPwocrMtqzh7O+5zIcVjtjnRRw9nF9X8OjYu+PrhZxFysV2J0NLDpaeOdV8urds5syvc8zlEue1K+wukCtjqw5VF8P7XHfd3CXiZcV6xytRV/Ou1TV0Xh+58aUJ08xxU/KWxm2LFrc7ga0xbSl3U9sPdBy8l9F56H7mg+yHuY8Odx1+fPjJ4e7cpznPsnoO9ab3pfbvfR73V9TA7sHIFzFDSS8PDh97VT7S8PrB6Ms3n8bBW+KE4Dv5SZ0p82m/mbPvP31Unk361Pr517zmQtyXy1/fLbIvWX5P+dHwc3qFe9VhLftX5/b8GyP0kTuQn1Ht6EMYR6w4dgF3E59BcKDhphmlPU8XTq/OgGBoZ0xnsiAyEHtJR5ltWRhYnrJms5mwQ+zNHBGcQpxDXDncOtyfeEp5zXi/8ZXxm/F/FjguqCE4IrRXmF+4VcRbZFW0WExJrFs8QHxV4piklGSblKPUlHSqjIjMkGzuDoMd3+Sq5D0V6BTaFCOVBJT6ldNUFFTGVfPUtNU+qZdqmGssap7Xstf6pV2n466L1b2hR9Yn6t81iDTkN+w1SjNWMp42KTG1he8dt82jLKQs3lmWWXlYs1o/tymwdbAj2Q3an3TwdhR2/OB01TnRxdiVwXXYrdI92sPAk9Zz0OvMzmBvBe+VXfd98n29/CT8lsid/scCfAMVg1BBg8G1lJQQp1DpMHTYm/BbEcW74yNdozSovNGo6NmYgdj2uPr4soS8PamJ8Umhyf57d+5zS3Ha75Bqn2Z/wCHd6aB7xs5DAZmhWdHZKTmZuQWHy47U5DUevZffVzB67PMJVKFEkVfx0ZP3S1ZOyZ72Kztx5vHZ1XKF8wEVJRd6qlDVWjXxtfV1Hy9JXg6pr70yd02lYf/17kaOprDmzlt8t1Na3rZatbW0y3dc7JS6f/WhwaPhxwndfE97ew73OT0XHQCDH4feDX94Dd6IjO+aqJ1CzyR+BJ8q5slfdZfUfjqvFm/M/9b3vo2CUQHgxCEANr7nONQAkHsRALE9ALDBuacdHQBO6gAhYAqgxQ4AWWj/OT8gOPEkABLgAZJADZjB+WUYnFMWgXrQBSbAGsQOKUOOUDR0AmqGXsM5nzTCBZGKqEeMIhmQBsh45CXkFJyleaFKUa/gTMwHfQH9CaOCScU8w/Jgw7FtOBKOgmvHc+Jj8f0ERUIRYZWGTPOUVp22mo6dLpceQZ9E/50hlmGJMZEJYsomshIrSGqkXuYQFizLBVZj1im2THYp9l6OGE4OzlYuf24a7ms8HrxI3kt8nnBG0CeQJ2grxCT0XLhIxFNUQHRK7KJ4tISGJCTZJZUv7QlH57xs/44WuUr5AoV9ihQlR2UNFT5VSHVMrUX9uEaIprYWvdaIdo1OjK6uHk6vT7/B4Lphk9Et41aTe6ZdZj3mgxajltNWC9Yrtjg7VntRBzVHKyeyc7JLsWur24wHyVPfK3JnufeAD8FX3y+J3Oz/PVAtKCm4PYQQ6hxWEb642yyyLGouWismJ3Y0Xinh6J6FJNfkB/u0U1pTLdMm0rMytDNBVl/OlcOn8gryzY4hj98vzC8OKDE8JV0meFakXKnCpjKqurT2ySVQr3rVpsH9RnBT8s0Tt6/d6W9b6uDtNHsQ8+js42fdaz0yfTufHxm4O0QaJo9cGp0d555Qm9Sbln9P/+HF7JHPO+baF8y+dH5TWCxZWv5h//PC8sKqxlrKr7ub+8fW/BPh+ZcAqsAEuIIQsA8cB3WgE4yCHxAJkoNsoAjoKNQAvUQAhASc5achriLewnm8FTId2YZcQWmjDqC60ezoQHQjBo/xxjRiGbFh2Cc4aVwebhHvhX9AkCUU0SBpomjGaZ1pH9MZ0rXSa9HfgbPYR4z2jKNwnrpOPE6SJT1ljoAzz2ZWXzYatmb2QA5Wjoece7ikuca5i3hseXG8HXz7+Q0EMAJPBQuFfIVlhVdFukXLxKLEjSW4JL5IPpY6L50i4ymruUNCjl0eL7+qMKc4ofRC+bHKbdWLaiXqhzSomp5ahtqSOow6i7rDeq369QZXDRuMmoxvm7SZdpo9Nu+1eGH5xmraesFmxQ5nz+og5qjuZO3s77LXtcTtpvuwx5qX4E4L75hdZ316/CCyin9EQE3gVLAoJSTkauhyuGlE4e6ZKC3q3ui2WFScVXxRwlSietKR5Ol9xinVqfRpew5Mw/tJb6ZF1sMcs9zuIw55Y/kpx3iP3y0MLKY/2Vzqf5pU9uDs3nKV818uXK2KrdGqw1wcuHzhSvI1r+sqjfRNEzev3z5wx6aN/e54R00n9YHWI2zX4JOap/t7vPp0nosMMA0+GnJ+Ofkq8TXz6LUxp/HViepJ92mGma4PmbOWnxnnXiyc/RqyqPId8aNnuXQ16Jfi9vwjAQbQbu4A4kAFjgA3EA4OgjPgNhiG178gZAHFQhXQIIIGYQSv/A4kDmmPPIP8grJAVaHxaCr6DcYJXu022AEcGfcTX0hQJ0zSnKTVox2hS6Lnp+9iiGeUZJxgOkP0I0mQvjM/ZCljTWLzZNfjkOJk56LhRnCv8izzrvIDASx8A+URlhXRFnUQCxLfL3FS8gacd8/LMu5QkHOV36dQodijtKIioequVqDer8ms5a5doTOnp61/2OCNkaJxjsm4mZZ5ocUXKzvrS7a0dmH2jx0lnXKcP7hauNV64D0pXg+9RXcd9Jn0MyRXBiAD/YPuUURDMkJnwq0i6iNZohKoYzFGsZfj2RP27fmY5AavU5WUqlSOtCPpqIPJGV8yPbKuZq/nOh2uOrJ81DH/8jHCccqJB0VSxbkn50pdT90pEz2TD+/9/ue7L2hWVlUz1STWTl10vNRSL3ol7+pSg/f1B40yTUeb52/Z3758h9Aa2NbaTuwIuNd4H/XA7mHpo4nHEk8o3ZVPx3s4e+37DvZff/52gDAo98JhiPryyHDNq3sjA6+nRhferI5Db7ETmHeYSTC5PPVpenTm6fvmD+UfM2cjPll/lprDzr2Zb17I+uLxVeLrl28ti2lLRt8x3zt/pPzU/LmwfGHFY5Ww2rhG/kX369q6+8b8RwcqKmweHxCtAQDo0fX1r6IAYAsAWMtfX18pX19fOw8nGyMA3A3b+g9p86xhBOAs1wbqu7L4X//l/B9aBdSDoBEc3gAAAD9QTFRFFwAHAQAATgAZEAAAMQAQYwAL/wB7/wD/qgBV/wAAfwB/rwAjmAAyAAAApQA2cAAlAAAAAAAAAAAAAAAAAAAAfJlyggAAABB0Uk5ToF70ItQDAgEDAQID/gH9/v65SD8AAAlGSURBVHjatVvZlqQgDA2LWjWN5f//7ZCFTUXAsnyY7jOt5nKzGEIC69BlFv6hFIC1VvNlLYBSRm4ZeyMM3DvRv1601tvm3Ce7nHOb1h6GSTc+D8CvDqzeXBQZLydwPAoLaggD9C9egSXhjsQg6crQ5X9BjWhmxf+VMEwPAphQ6yTdOVyiOtezUUKQ2yyY5xhA8aCdrK324jmYJzAGpuEBAEm8vHIRSYaIZ1VEWDObqd6IBn//+3sGSHwgdY5UkyPItXkPsMEFGAMqjCFMXwIQ8bj4P6FYBx8U+cET0RFVgKn4waYtNACoKH6ek4I/7vT6iJGizFegTsM1CRcA3quxXpeRx2BelxeC0FYF40FFWHWXAVp+FG/b0gMIdBdDjylcgoZ7AECefYkl9EkXHhxBMGJD1lTVADXnU1aWb4bFCw0IwVsO6tFpVXNIqMnXXvtA4gnKjYvEMn5+1wgDsIXHDWrC3bvIgP/90RJqCKAq3+vNiCHevtgL/TpIn/0A0u2I5KuLFpKW1AdACDO3tX+wBPOuIoCa/LdBQ3QPXOjL//htJwj2AOYVSL6Zvqc/2SLzeYoA9v4n8ufn5KMagCPriSWWAAyKFfnuMfmObdpzsCGU6YoBvsVMz8qPCGR5dQBMkreYh+UzgoURqCoDy2oTUueeRuDJN+hiusYArZv+qp6XzwiMrHE+Z0AJPw/5/5k3sjPmZpAzIH8yv5GPEUmxK+RmACkCePXY9YU/q9lWnor2JyfxGf/+2bCc+ciAR0YIqw4gO2E9Jpxzdgt8JaanHYAZrQMkWJzzZ1Ko6DU72sgdoi8u1ewYoMjjFfBCHOcXCG1v25ua+uzYpF2UMSLphUqOShAAi+dlYwXUTDhiVj1BSrJS3tFr2sPrsHVHT4h2CDkBU90DP9kep+0lIRPyGYXOKxmcrxuWNmcAhIC6B+DlY2kA27pk9bxD3OXrfp0iLgGgdxIBVwYWjGBpBIq4F4LtRFnsAUTBkhhgRBcWyABMzFla6cckiei5Oy1zogAKC7ggIA9fl/fpYGgXKxHK35GBDgIy112uTMUq3krW1YQUoBr5hSArQqO8tICP3uUtFfkmJhbukgLLPxDAwkGwYQH+jn9F4nAl31y+yq+FVqsDA0TH0oqxRSpVuZflvy+9mc1pXsUMIYUFaD20y51q8jsSSrv+oTiyRn6Z8j/0NW3mkL3e/lzRnaIDEA20Hits8JSC4H/L2vG1Avr8IqsQfQDaj2CZqkZByLd7InVUudcBTPzL0vrEULK2gjbvlD/s9Dp1eGCpA0sMsDk2NEA2iKhjLlM+EAxgbnhA+CAoCn/+Ib8wMYY2ZNw5R0srJSUn6UyYohFAMIGG6+r1kM4VoiwV0zoJoPctrHvgH2vbBKj+urlTCoYJoASLlg7iDa0HxQZDJN8LEwvscoGIGP3fA9CMxrVtkF8fV5ukJQL6NzUgVggEo98GebnzTlzKcbu3DJYi1qaAFLE0AWQCDxTE/+gKgjsrRADQfjLY4CfGMRbIiPS6Dm9ahHa/W+p0ArRBdfgysumIWZoBDYjhOSsAro0nt0GhYEkB2f91CRW+Tz8Adj0CgB6xfbptcO+KSQOmv7BBwZgA2PBrx7c40kQZVIxGCc1IZYmXraErDhU2WFJgNneIC72BYASAKReYu6Ie2DTuAOgE4Iq8vQ2WsWdV63AYHAVwsMF9kjoaBU4AdOxJygV+imLbYBQYBCB07+/ZVX2nGwDICF3bDdEGD++PrtizY7xww3YgCuo+hLl94duOnCzmkbARiqU8dMC436sMeeEewNXDZIOnJmZLCsYA8McIOB+4fhg/PWcf7L0rDgGIn2MBAA0bXM6sZOeKgwB8QuK/Q+2UjNZ57mM7CoYCYUzJmkmp2ODp22NicOtTwEmppOV1P+Q4WAsUuSsOuGGQqXs2JrC+qxuezBWHImG+MWlszXgfWY+UKTEYApBtzRpWiGs09SjbWz682Jy29matCmZyxf58gGjlykyrQNGsYScKBr5GZYHiskTTUcFMrtitg1SiMY0iVdjFuk4K+remi1RmIBRK53Md2I6aQ5YedlYHyjIdRsGKDijpaJUdEwVTrxWkIlko1WJcODLdVbooPsumRwn42iWVatcpgnHH0l9P1a2IBV1FOjZsm5fr/f9NuzOOXvnlZ/m87+eYzsoxpRxYWFkExIa1Dzcy9Z5jZ6X006OarTSsOR1eglRWxJupZY+eoIMv6JRf9kVQ69mnPMTLDQmS1stDq9A2SI2itJb+gkeRmhhsAE7H1pvNv/ZIwFIeWoVTq7noBuaGyP7va+qcnKTv1Iau7+PGGmL0KA4u6elFGrgHm8i4R2g67S+f3gkAnlCv+4PLnALuXLebu9dAuGKHAL/m5Rc2T9NfESLjCXV5dMunocAWhA27tzr4at3MCulMQatsosiP733cRf/83O4f5G7S1DFgaAyAuoCz2C4x4N9pA8P0bQcL9RdvZH2aju1dzmZQwLGBIfbQPNFD9NlNHuzrjactHFEJZnCXPYTMW95UbWJZDPUYmcXYn7bxQDxfO21kCr1Gv5C/qdQuV2/lUj9r5dqCAcjxzpGBWVrNftbMZubLZjbJ7H/TTkfn88TxRTtfbCicH0eQtRNeNjRy/y88jgDf+Tby8suWTvMTBNJWS63FqQ3inIGXtNU+6Y34viXIbzY28z7gscbqEH/eFfmnrd0QWsvNI63dVlH2ey7/vLk9NdfjaMD36i/65bvb+z1xf9OXaohDCoPt/QutXEMaErn5Wd4gDimMDThM/AyrgYZ97mpf1F+ftKn8v7AmQyJqHMKHWwrDiIpWtTGbxpjPZuO8kBsa89loyTOPqGzW3B90khcNDDqFQTv86KZX3AAwyYSN5v4V3Cr1jHqFicuXqO5q+R3DbtbJtBklEbx3reXt2bznvAbL0eqbabtJtro8bbjwno2GS8vUN8zdStvsHPtJ9bfjfutu3nIJW53QrMuzt5padlUcujQ8c6k7Jk97hl7DxKlMsi7ZvkcuEyMY7S1tgjw9AGBK9hdbk9d1qRgNWC5P0Orf6yMMyLyhdnG02VxO/Yam7q7J487BZyPUbqFhXiafPfs89kw2UQx9Pjn4LBOgYej0MPudyjFXk8lPzJ7vhn7zDeiWOcL0QwChBsO0ywXZ9P88OP7/H8UW0gbr5dFEAAAAAElFTkSuQmCC";

/* ================================================================== */
/* Modello                                                             */
/* ================================================================== */

const KEY = "crm:data:v6";
const KEY_V5 = "crm:data:v5";

const PIATTAFORME = [
  { v: "meta", l: "Meta" }, { v: "google", l: "Google Ads" }, { v: "tiktok", l: "TikTok" },
  { v: "linkedin", l: "LinkedIn" }, { v: "youtube", l: "YouTube" }, { v: "amazon", l: "Amazon Ads" }, { v: "altro", l: "Altro" },
];
const OBIETTIVI = [
  { v: "vendite", l: "Vendite" }, { v: "lead", l: "Lead" }, { v: "traffico", l: "Traffico" }, { v: "awareness", l: "Notorietà" },
];
const BMK_DEFAULT = {
  vendite: { roasMin: 2.5, cpaMax: 0, cplMax: 0, ctrMin: 1, freqMax: 3.5, cpmMax: 0 },
  lead: { roasMin: 0, cpaMax: 35, cplMax: 35, ctrMin: 1, freqMax: 3, cpmMax: 0 },
  traffico: { roasMin: 0, cpaMax: 0, cplMax: 0, ctrMin: 1.5, freqMax: 2.5, cpmMax: 12 },
  awareness: { roasMin: 0, cpaMax: 0, cplMax: 0, ctrMin: 0.6, freqMax: 2.5, cpmMax: 8 },
};
const BMK_VUOTO = { roasMin: 0, cpaMax: 0, cplMax: 0, ctrMin: 0, freqMax: 0, cpmMax: 0 };
/* Nelle campagne a contatti il lead è il primo punto di conversione, e il costo
   che conta è il CPL. Finché nessuno l'ha impostato ripiego sul CPA, così le
   campagne scritte prima che questo campo esistesse continuano a essere giudicate. */
const tettoCpl = (b) => ((b || {}).cplMax > 0 ? b.cplMax : (b || {}).cpaMax || 0);
/* Una campagna si misura sui lead se l'obiettivo è quello o se lead ne sono arrivati */
const suiLead = (k, d) => k.obiettivo === "lead" || ((d || {}).ld || 0) > 0;
const kpiCosto = (k, d) => (suiLead(k, d) ? tettoCpl(k.bmk) : (k.bmk || {}).cpaMax || 0);
const STATI_CLIENTE = [
  { v: "attivo", l: "Attivo" }, { v: "onboarding", l: "Onboarding" }, { v: "pausa", l: "In pausa" }, { v: "chiuso", l: "Chiuso" },
];
/* ---------- lingue ----------
   L'interfaccia è tradotta; i testi lunghi — avvisi, guida, consigli del semaforo —
   restano in italiano finché non vengono tradotti anch'essi, lingua per lingua.
   La chiave manca? si ricade sull'italiano, che resta la lingua di riferimento. */
/* Ogni lingua porta il suo formato e la valuta che ci si aspetta di leggerci dentro.
   L'inglese è due lingue diverse: formati e valuta non coincidono. */
const LINGUE = [
  { v: "it", l: "Italiano", b: "IT", loc: "it-IT", val: "EUR" },
  { v: "en", l: "English (US)", b: "US", loc: "en-US", val: "USD" },
  { v: "en-gb", l: "English (UK)", b: "GB", loc: "en-GB", val: "GBP" },
  { v: "es", l: "Español", b: "ES", loc: "es-ES", val: "EUR" },
  { v: "fr", l: "Français", b: "FR", loc: "fr-FR", val: "EUR" },
  { v: "de", l: "Deutsch", b: "DE", loc: "de-DE", val: "EUR" },
];

/* Valute: il simbolo e la posizione li mette Intl, qui serve solo il codice */
const VALUTE = ["EUR", "USD", "GBP", "CHF", "CAD", "AUD", "SEK", "NOK", "DKK", "PLN", "CZK", "RON", "BRL", "MXN", "ARS", "JPY", "INR", "AED"];

const DIZ = {
  it: {
    "a.predefinita": "dall'agenzia",
    "a.valuta": "Valuta",
    "a.lingua": "Lingua",
    /* etichette degli elenchi */
    "e.attiva": "Attiva", "e.test": "In test", "e.pausa": "In pausa", "e.conclusa": "Conclusa",
    "e.errorePag": "Errore pagamento", "e.onboarding": "Onboarding", "e.chiuso": "Chiuso",
    "e.adset": "Ad set", "e.annuncio": "Annuncio", "e.vendite": "Vendite", "e.lead": "Lead",
    "e.traffico": "Traffico", "e.notorieta": "Notorietà", "e.urgente": "Urgente", "e.alta": "Alta",
    "e.media": "Media", "e.bassa": "Bassa", "e.dafare": "Da fare", "e.incorso": "In corso",
    "e.attesa": "In attesa", "e.fatto": "Fatto",
    /* menu */
    "nav.cruscotto": "Panoramica", "nav.clienti": "Clienti", "nav.sottoclienti": "Sottoclienti",
    "nav.advertising": "Advertising", "nav.mesi": "Riepilogo mensile", "nav.semaforo": "Semaforo",
    "nav.pixel": "Pixel facile", "nav.calendario": "Calendario", "nav.avvisi": "Avvisi",
    "nav.grafici": "Grafici", "nav.vendite": "Ciclo di vendita", "nav.chiusure": "Chiusure",
    "nav.pianificatore": "Pianificatore", "nav.funnel": "Funnel", "nav.diario": "Diario",
    "nav.file": "File", "nav.economia": "Economia", "nav.ai": "Consulente", "nav.task": "Task",
    "nav.automazioni": "Automazioni", "nav.importa": "Importa dati", "nav.backup": "Backup",
    "nav.guida": "Guida", "nav.impostazioni": "Impostazioni", "nav.esci": "Esci", "nav.contatti": "Contatti",
    "nav.pipeline": "Pipeline", "nav.messaggio": "Messaggio", "nav.report": "Report", "nav.winner": "Winner ads", "nav.academy": "Academy",
    "nav.scheda": "Scheda cliente", "nav.campagne": "Advertising",
    /* azioni ricorrenti */
    "a.salva": "Salva", "a.annulla": "Annulla", "a.elimina": "Elimina", "a.aggiungi": "Aggiungi",
    "a.chiudi": "Chiudi", "a.apri": "Apri", "a.copia": "Copia", "a.copiato": "Copiato",
    "a.importa": "Importa", "a.esporta": "Esporta", "a.azzera": "Azzera filtri", "a.oggi": "Oggi",
    "a.archivia": "Archivia", "a.modifica": "Modifica", "a.conferma": "Conferma", "a.indietro": "Indietro",
    "a.tutti": "Tutti", "a.nessuno": "Nessuno", "a.filtri": "Filtri", "a.cerca": "Cerca",
    /* intestazioni di colonna */
    "c.campagna": "Campagna", "c.cliente": "Cliente", "c.stato": "Stato", "c.fase": "Fase",
    "c.speso": "Speso", "c.spesa": "Spesa", "c.lead": "Lead", "c.cpl": "CPL", "c.cpa": "CPA",
    "c.roas": "ROAS", "c.ctr": "CTR link", "c.cpc": "CPC link", "c.conversioni": "Conv.",
    "c.avvisi": "Avvisi", "c.resa": "Resa", "c.campagne": "Campagne", "c.budget": "Budget",
    "c.contratti": "Contratti", "c.incassato": "Incassato", "c.data": "Data", "c.account": "Account",
    "c.livello": "Livello", "c.risultati": "Risultati", "c.costo": "Costo", "c.quota": "Quota",
    /* parole di stato */
    "s.attiva": "Attiva", "s.test": "In test", "s.pausa": "In pausa", "s.conclusa": "Conclusa",
    "s.errore": "Errore pagamento", "s.nessunDato": "nessun dato", "s.tutti": "Tutti i clienti",
    /* semaforo */
    "sem.verde": "Scala", "sem.giallo": "Non fare nulla", "sem.rosso": "Spegni",
    "sem.grigio": "Non ancora valutabile", "sem.valutazione": "in valutazione dopo modifica",
    "sem.pronta": "valutazione pronta",
  },

  en: {
    "a.predefinita": "from agency",
    "a.valuta": "Currency",
    "a.lingua": "Language",
    "e.attiva": "Active", "e.test": "Testing", "e.pausa": "Paused", "e.conclusa": "Ended",
    "e.errorePag": "Payment error", "e.onboarding": "Onboarding", "e.chiuso": "Closed",
    "e.adset": "Ad set", "e.annuncio": "Ad", "e.vendite": "Sales", "e.lead": "Leads",
    "e.traffico": "Traffic", "e.notorieta": "Awareness", "e.urgente": "Urgent", "e.alta": "High",
    "e.media": "Medium", "e.bassa": "Low", "e.dafare": "To do", "e.incorso": "In progress",
    "e.attesa": "Waiting", "e.fatto": "Done",
    "nav.cruscotto": "Overview", "nav.clienti": "Clients", "nav.sottoclienti": "Sub-clients",
    "nav.advertising": "Advertising", "nav.mesi": "Monthly summary", "nav.semaforo": "Traffic light",
    "nav.pixel": "Easy Pixel", "nav.calendario": "Calendar", "nav.avvisi": "Alerts",
    "nav.grafici": "Charts", "nav.vendite": "Sales cycle", "nav.chiusure": "Closed deals",
    "nav.pianificatore": "Planner", "nav.funnel": "Funnel", "nav.diario": "Journal",
    "nav.file": "Files", "nav.economia": "Finances", "nav.ai": "Advisor", "nav.task": "Tasks",
    "nav.automazioni": "Automations", "nav.importa": "Import data", "nav.backup": "Backup",
    "nav.guida": "Guide", "nav.impostazioni": "Settings", "nav.esci": "Log out", "nav.contatti": "Contacts",
    "nav.pipeline": "Pipeline", "nav.messaggio": "Messaging", "nav.report": "Report", "nav.winner": "Winner ads", "nav.academy": "Academy",
    "nav.scheda": "Client record", "nav.campagne": "Advertising",
    "a.salva": "Save", "a.annulla": "Cancel", "a.elimina": "Delete", "a.aggiungi": "Add",
    "a.chiudi": "Close", "a.apri": "Open", "a.copia": "Copy", "a.copiato": "Copied",
    "a.importa": "Import", "a.esporta": "Export", "a.azzera": "Clear filters", "a.oggi": "Today",
    "a.archivia": "Archive", "a.modifica": "Edit", "a.conferma": "Confirm", "a.indietro": "Back",
    "a.tutti": "All", "a.nessuno": "None", "a.filtri": "Filters", "a.cerca": "Search",
    "c.campagna": "Campaign", "c.cliente": "Client", "c.stato": "Status", "c.fase": "Phase",
    "c.speso": "Spent", "c.spesa": "Spend", "c.lead": "Leads", "c.cpl": "CPL", "c.cpa": "CPA",
    "c.roas": "ROAS", "c.ctr": "Link CTR", "c.cpc": "Link CPC", "c.conversioni": "Conv.",
    "c.avvisi": "Alerts", "c.resa": "Return", "c.campagne": "Campaigns", "c.budget": "Budget",
    "c.contratti": "Contracts", "c.incassato": "Collected", "c.data": "Date", "c.account": "Account",
    "c.livello": "Level", "c.risultati": "Results", "c.costo": "Cost", "c.quota": "Share",
    "s.attiva": "Active", "s.test": "Testing", "s.pausa": "Paused", "s.conclusa": "Ended",
    "s.errore": "Payment error", "s.nessunDato": "no data", "s.tutti": "All clients",
    "sem.verde": "Scale", "sem.giallo": "Do nothing", "sem.rosso": "Turn off",
    "sem.grigio": "Not yet assessable", "sem.valutazione": "under review after change",
    "sem.pronta": "review ready",
  },

  es: {
    "a.predefinita": "de la agencia",
    "a.valuta": "Moneda",
    "a.lingua": "Idioma",
    "e.attiva": "Activa", "e.test": "En prueba", "e.pausa": "En pausa", "e.conclusa": "Finalizada",
    "e.errorePag": "Error de pago", "e.onboarding": "Alta", "e.chiuso": "Cerrado",
    "e.adset": "Conjunto de anuncios", "e.annuncio": "Anuncio", "e.vendite": "Ventas", "e.lead": "Leads",
    "e.traffico": "Tráfico", "e.notorieta": "Notoriedad", "e.urgente": "Urgente", "e.alta": "Alta",
    "e.media": "Media", "e.bassa": "Baja", "e.dafare": "Por hacer", "e.incorso": "En curso",
    "e.attesa": "En espera", "e.fatto": "Hecho",
    "nav.cruscotto": "Panorama", "nav.clienti": "Clientes", "nav.sottoclienti": "Subclientes",
    "nav.advertising": "Publicidad", "nav.mesi": "Resumen mensual", "nav.semaforo": "Semáforo",
    "nav.pixel": "Píxel fácil", "nav.calendario": "Calendario", "nav.avvisi": "Avisos",
    "nav.grafici": "Gráficos", "nav.vendite": "Ciclo de venta", "nav.chiusure": "Cierres",
    "nav.pianificatore": "Planificador", "nav.funnel": "Embudo", "nav.diario": "Diario",
    "nav.file": "Archivos", "nav.economia": "Economía", "nav.ai": "Asesor", "nav.task": "Tareas",
    "nav.automazioni": "Automatizaciones", "nav.importa": "Importar datos", "nav.backup": "Copia de seguridad",
    "nav.guida": "Guía", "nav.impostazioni": "Ajustes", "nav.esci": "Cerrar sesión", "nav.contatti": "Contactos",
    "nav.pipeline": "Pipeline", "nav.messaggio": "Mensaje", "nav.report": "Informe",
    "nav.scheda": "Ficha del cliente", "nav.campagne": "Publicidad",
    "a.salva": "Guardar", "a.annulla": "Cancelar", "a.elimina": "Eliminar", "a.aggiungi": "Añadir",
    "a.chiudi": "Cerrar", "a.apri": "Abrir", "a.copia": "Copiar", "a.copiato": "Copiado",
    "a.importa": "Importar", "a.esporta": "Exportar", "a.azzera": "Borrar filtros", "a.oggi": "Hoy",
    "a.archivia": "Archivar", "a.modifica": "Editar", "a.conferma": "Confirmar", "a.indietro": "Atrás",
    "a.tutti": "Todos", "a.nessuno": "Ninguno", "a.filtri": "Filtros", "a.cerca": "Buscar",
    "c.campagna": "Campaña", "c.cliente": "Cliente", "c.stato": "Estado", "c.fase": "Fase",
    "c.speso": "Gastado", "c.spesa": "Gasto", "c.lead": "Leads", "c.cpl": "CPL", "c.cpa": "CPA",
    "c.roas": "ROAS", "c.ctr": "CTR de enlace", "c.cpc": "CPC de enlace", "c.conversioni": "Conv.",
    "c.avvisi": "Avisos", "c.resa": "Rendimiento", "c.campagne": "Campañas", "c.budget": "Presupuesto",
    "c.contratti": "Contratos", "c.incassato": "Cobrado", "c.data": "Fecha", "c.account": "Cuenta",
    "c.livello": "Nivel", "c.risultati": "Resultados", "c.costo": "Coste", "c.quota": "Cuota",
    "s.attiva": "Activa", "s.test": "En prueba", "s.pausa": "En pausa", "s.conclusa": "Finalizada",
    "s.errore": "Error de pago", "s.nessunDato": "sin datos", "s.tutti": "Todos los clientes",
    "sem.verde": "Escala", "sem.giallo": "No hagas nada", "sem.rosso": "Apaga",
    "sem.grigio": "Aún no evaluable", "sem.valutazione": "en evaluación tras el cambio",
    "sem.pronta": "evaluación lista",
  },

  fr: {
    "a.predefinita": "de l'agence",
    "a.valuta": "Devise",
    "a.lingua": "Langue",
    "e.attiva": "Active", "e.test": "En test", "e.pausa": "En pause", "e.conclusa": "Terminée",
    "e.errorePag": "Erreur de paiement", "e.onboarding": "Intégration", "e.chiuso": "Clôturé",
    "e.adset": "Ensemble de publicités", "e.annuncio": "Publicité", "e.vendite": "Ventes", "e.lead": "Leads",
    "e.traffico": "Trafic", "e.notorieta": "Notoriété", "e.urgente": "Urgent", "e.alta": "Haute",
    "e.media": "Moyenne", "e.bassa": "Basse", "e.dafare": "À faire", "e.incorso": "En cours",
    "e.attesa": "En attente", "e.fatto": "Terminé",
    "nav.cruscotto": "Vue d'ensemble", "nav.clienti": "Clients", "nav.sottoclienti": "Sous-clients",
    "nav.advertising": "Publicité", "nav.mesi": "Récapitulatif mensuel", "nav.semaforo": "Feu tricolore",
    "nav.pixel": "Pixel facile", "nav.calendario": "Calendrier", "nav.avvisi": "Alertes",
    "nav.grafici": "Graphiques", "nav.vendite": "Cycle de vente", "nav.chiusure": "Ventes conclues",
    "nav.pianificatore": "Planificateur", "nav.funnel": "Tunnel", "nav.diario": "Journal",
    "nav.file": "Fichiers", "nav.economia": "Finances", "nav.ai": "Conseiller", "nav.task": "Tâches",
    "nav.automazioni": "Automatisations", "nav.importa": "Importer", "nav.backup": "Sauvegarde",
    "nav.guida": "Guide", "nav.impostazioni": "Réglages", "nav.esci": "Déconnexion", "nav.contatti": "Contacts",
    "nav.pipeline": "Pipeline", "nav.messaggio": "Message", "nav.report": "Rapport",
    "nav.scheda": "Fiche client", "nav.campagne": "Publicité",
    "a.salva": "Enregistrer", "a.annulla": "Annuler", "a.elimina": "Supprimer", "a.aggiungi": "Ajouter",
    "a.chiudi": "Fermer", "a.apri": "Ouvrir", "a.copia": "Copier", "a.copiato": "Copié",
    "a.importa": "Importer", "a.esporta": "Exporter", "a.azzera": "Effacer les filtres", "a.oggi": "Aujourd'hui",
    "a.archivia": "Archiver", "a.modifica": "Modifier", "a.conferma": "Confirmer", "a.indietro": "Retour",
    "a.tutti": "Tous", "a.nessuno": "Aucun", "a.filtri": "Filtres", "a.cerca": "Rechercher",
    "c.campagna": "Campagne", "c.cliente": "Client", "c.stato": "Statut", "c.fase": "Phase",
    "c.speso": "Dépensé", "c.spesa": "Dépense", "c.lead": "Leads", "c.cpl": "CPL", "c.cpa": "CPA",
    "c.roas": "ROAS", "c.ctr": "CTR lien", "c.cpc": "CPC lien", "c.conversioni": "Conv.",
    "c.avvisi": "Alertes", "c.resa": "Rendement", "c.campagne": "Campagnes", "c.budget": "Budget",
    "c.contratti": "Contrats", "c.incassato": "Encaissé", "c.data": "Date", "c.account": "Compte",
    "c.livello": "Niveau", "c.risultati": "Résultats", "c.costo": "Coût", "c.quota": "Part",
    "s.attiva": "Active", "s.test": "En test", "s.pausa": "En pause", "s.conclusa": "Terminée",
    "s.errore": "Erreur de paiement", "s.nessunDato": "aucune donnée", "s.tutti": "Tous les clients",
    "sem.verde": "Augmenter", "sem.giallo": "Ne rien faire", "sem.rosso": "Désactiver",
    "sem.grigio": "Pas encore évaluable", "sem.valutazione": "en évaluation après modification",
    "sem.pronta": "évaluation prête",
  },

  de: {
    "a.predefinita": "von der Agentur",
    "a.valuta": "Währung",
    "a.lingua": "Sprache",
    "e.attiva": "Aktiv", "e.test": "Im Test", "e.pausa": "Pausiert", "e.conclusa": "Beendet",
    "e.errorePag": "Zahlungsfehler", "e.onboarding": "Onboarding", "e.chiuso": "Geschlossen",
    "e.adset": "Anzeigengruppe", "e.annuncio": "Anzeige", "e.vendite": "Verkäufe", "e.lead": "Leads",
    "e.traffico": "Traffic", "e.notorieta": "Bekanntheit", "e.urgente": "Dringend", "e.alta": "Hoch",
    "e.media": "Mittel", "e.bassa": "Niedrig", "e.dafare": "Zu erledigen", "e.incorso": "In Arbeit",
    "e.attesa": "Wartend", "e.fatto": "Erledigt",
    "nav.cruscotto": "Übersicht", "nav.clienti": "Kunden", "nav.sottoclienti": "Unterkunden",
    "nav.advertising": "Werbung", "nav.mesi": "Monatsübersicht", "nav.semaforo": "Ampel",
    "nav.pixel": "Pixel leicht", "nav.calendario": "Kalender", "nav.avvisi": "Hinweise",
    "nav.grafici": "Diagramme", "nav.vendite": "Verkaufszyklus", "nav.chiusure": "Abschlüsse",
    "nav.pianificatore": "Planer", "nav.funnel": "Funnel", "nav.diario": "Journal",
    "nav.file": "Dateien", "nav.economia": "Finanzen", "nav.ai": "Berater", "nav.task": "Aufgaben",
    "nav.automazioni": "Automatisierungen", "nav.importa": "Daten importieren", "nav.backup": "Sicherung",
    "nav.guida": "Handbuch", "nav.impostazioni": "Einstellungen", "nav.esci": "Abmelden", "nav.contatti": "Kontakte",
    "nav.pipeline": "Pipeline", "nav.messaggio": "Botschaft", "nav.report": "Bericht",
    "nav.scheda": "Kundenblatt", "nav.campagne": "Werbung",
    "a.salva": "Speichern", "a.annulla": "Abbrechen", "a.elimina": "Löschen", "a.aggiungi": "Hinzufügen",
    "a.chiudi": "Schließen", "a.apri": "Öffnen", "a.copia": "Kopieren", "a.copiato": "Kopiert",
    "a.importa": "Importieren", "a.esporta": "Exportieren", "a.azzera": "Filter zurücksetzen", "a.oggi": "Heute",
    "a.archivia": "Archivieren", "a.modifica": "Bearbeiten", "a.conferma": "Bestätigen", "a.indietro": "Zurück",
    "a.tutti": "Alle", "a.nessuno": "Keine", "a.filtri": "Filter", "a.cerca": "Suchen",
    "c.campagna": "Kampagne", "c.cliente": "Kunde", "c.stato": "Status", "c.fase": "Phase",
    "c.speso": "Ausgegeben", "c.spesa": "Ausgaben", "c.lead": "Leads", "c.cpl": "CPL", "c.cpa": "CPA",
    "c.roas": "ROAS", "c.ctr": "Link-CTR", "c.cpc": "Link-CPC", "c.conversioni": "Konv.",
    "c.avvisi": "Hinweise", "c.resa": "Ertrag", "c.campagne": "Kampagnen", "c.budget": "Budget",
    "c.contratti": "Verträge", "c.incassato": "Vereinnahmt", "c.data": "Datum", "c.account": "Konto",
    "c.livello": "Ebene", "c.risultati": "Ergebnisse", "c.costo": "Kosten", "c.quota": "Anteil",
    "s.attiva": "Aktiv", "s.test": "Im Test", "s.pausa": "Pausiert", "s.conclusa": "Beendet",
    "s.errore": "Zahlungsfehler", "s.nessunDato": "keine Daten", "s.tutti": "Alle Kunden",
    "sem.verde": "Skalieren", "sem.giallo": "Nichts tun", "sem.rosso": "Abschalten",
    "sem.grigio": "Noch nicht bewertbar", "sem.valutazione": "in Prüfung nach Änderung",
    "sem.pronta": "Prüfung bereit",
  },
};

/* Campagna ferma ("errore"): la causa arriva da Meta con lo stato dell'account */
const CAUSE_FERMO = {
  pagamento: { l: "Errore pagamento", titolo: "Campagna ferma per errore di pagamento", causa: "Carta rifiutata, plafond superato o metodo di pagamento scaduto nel gestore inserzioni.", azione: "Sistema il metodo di pagamento nella piattaforma, poi riporta lo stato su Attiva: l'avviso sparisce e il fermo resta registrato nel Diario." },
  disattivato: { l: "Account disattivato", titolo: "Campagna ferma: account pubblicitario disattivato da Meta", causa: "Meta ha disattivato l'account pubblicitario: le campagne non erogano finché non viene riattivato.", azione: "Apri Qualità dell'account in Meta Business Suite e chiedi la revisione. Se l'account non torna, sposta le campagne su un altro account collegato al cliente." },
  revisione: { l: "Account in revisione", titolo: "Campagna ferma: account pubblicitario in revisione", causa: "Meta sta verificando l'account per un controllo di rischio: nel frattempo le campagne non erogano.", azione: "Controlla Qualità dell'account in Meta Business Suite: di solito basta attendere o completare la verifica richiesta." },
  chiuso: { l: "Account chiuso", titolo: "Campagna ferma: account pubblicitario chiuso", causa: "L'account pubblicitario è chiuso o in chiusura: non può più erogare.", azione: "Collega al cliente l'account nuovo e segna la campagna come conclusa." },
};
const causaFermo = (k) => CAUSE_FERMO[k.causaErrore] || CAUSE_FERMO.pagamento;
const etStato = (k) => (k.stato === "errore" ? causaFermo(k).l : et(STATI_CAMPAGNA, k.stato));

const STATI_CAMPAGNA = [
  { v: "attiva", l: "Attiva" }, { v: "test", l: "In test" }, { v: "errore", l: "Errore pagamento" },
  { v: "pausa", l: "In pausa" }, { v: "conclusa", l: "Conclusa" },
];
const TIPI_CONTATTO = [{ v: "lead", l: "Lead" }, { v: "referente", l: "Referente cliente" }];
const STATI_LEAD = [
  { v: "nuovo", l: "Nuovo" }, { v: "contattato", l: "Contattato" }, { v: "qualificato", l: "Qualificato" },
  { v: "opportunita", l: "Opportunità" }, { v: "cliente", l: "Cliente" }, { v: "scartato", l: "Scartato" },
];
const FASI = [
  { v: "nuovo", l: "Nuovo", p: 0.1 }, { v: "qualificato", l: "Qualificato", p: 0.3 },
  { v: "proposta", l: "Proposta", p: 0.55 }, { v: "negoziazione", l: "Negoziazione", p: 0.75 },
  { v: "vinto", l: "Vinto", p: 1 }, { v: "perso", l: "Perso", p: 0 },
];
const FASI_APERTE = FASI.slice(0, 4);
const faseDi = (v) => FASI.find((f) => f.v === v) || FASI[0];

const STATI_TASK = [
  { v: "dafare", l: "Da fare" }, { v: "incorso", l: "In corso" }, { v: "attesa", l: "In attesa" }, { v: "fatto", l: "Fatto" },
];
const PRIORITA = [
  { v: "urgente", l: "Urgente", r: 0 }, { v: "alta", l: "Alta", r: 1 }, { v: "media", l: "Media", r: 2 }, { v: "bassa", l: "Bassa", r: 3 },
];
const prioR = (v) => (PRIORITA.find((p) => p.v === v) || { r: 2 }).r;
const RICORRENZE = [
  { v: "no", l: "Nessuna", g: 0 }, { v: "7", l: "Ogni settimana", g: 7 }, { v: "14", l: "Ogni due settimane", g: 14 },
  { v: "30", l: "Ogni mese", g: 30 }, { v: "90", l: "Ogni trimestre", g: 90 },
];
const AREE_DEFAULT = [
  { id: "ar1", nome: "Advertising" }, { id: "ar2", nome: "Creatività" }, { id: "ar3", nome: "Reportistica" },
  { id: "ar4", nome: "Sito e tracciamento" }, { id: "ar5", nome: "Commerciale" }, { id: "ar6", nome: "Amministrazione" },
];

/* Percorso commerciale: ogni riga è un periodo (mese) per un cliente */
const FASI_VENDITA = [
  { k: "lead", l: "Lead" },
  { k: "chiamate", l: "Chiamate" },
  { k: "discoveryFissate", l: "Discovery fissate" },
  { k: "discoverySvolte", l: "Discovery svolte" },
  { k: "demoFissate", l: "Demo fissate" },
  { k: "demoSvolte", l: "Demo svolte" },
  { k: "chiusure", l: "Chiusure" },
];
const CAMPI_VENDITA = [...FASI_VENDITA.map((f) => f.k), "offerte", "valoreContratto", "incassato", "spesa"];
const FUNNEL_PROV = [
  { v: "vsl", l: "VSL / video" }, { v: "webinar", l: "Webinar" }, { v: "form", l: "Form contatti" },
  { v: "dm", l: "DM / messaggi" }, { v: "chiamata", l: "Chiamata a freddo" }, { v: "referral", l: "Passaparola" },
  { v: "organico", l: "Organico" }, { v: "evento", l: "Evento" }, { v: "altro", l: "Altro" },
];

/* --- automazioni --- */
/* Pacchetto ADA: una regola per ciascuna soglia del framework decisionale.
   Si appoggiano agli avvisi, che già conoscono i benchmark della singola campagna:
   così la soglia resta quella del cliente e non un numero scritto qui. */
const REGOLE_ADA = [
  { nome: "ADA · Gancio debole (CTR link)", trigger: "avviso", filtroAvviso: "CTR", soloCritici: "no",
    taskTitolo: "CTR link sotto soglia: il gancio non funziona", taskPriorita: "alta", taskGiorni: 2,
    perche: "Prima soglia del framework: se il CTR sul link è sotto l'obiettivo della campagna, il problema è l'annuncio, non il pubblico." },
  { nome: "ADA · Asta cara (CPM)", trigger: "avviso", filtroAvviso: "CPM", soloCritici: "no",
    taskTitolo: "CPM sopra la media dell'account", taskPriorita: "media", taskGiorni: 3,
    perche: "Seconda soglia: paghi l'asta più degli altri tuoi annunci. Di solito è qualità del creativo o pubblico troppo conteso." },
  { nome: "ADA · Spesa senza conversioni", trigger: "avviso", filtroAvviso: "Conversioni", soloCritici: "no",
    taskTitolo: "Spesa oltre soglia senza conversioni", taskPriorita: "urgente", taskGiorni: 1,
    perche: "Terza soglia: superata la spesa minima di valutazione senza un risultato, la campagna va spenta o rifatta." },
  { nome: "ADA · Sotto il pareggio (margine)", trigger: "avviso", filtroAvviso: "Margine", soloCritici: "no",
    taskTitolo: "ROAS sotto il punto di pareggio del cliente", taskPriorita: "urgente", taskGiorni: 1,
    perche: "Quarta soglia: il ritorno non copre il margine del cliente. È l'unica che misura soldi veri, non rendimento pubblicitario." },
  { nome: "ADA · Ritorno sotto obiettivo (ROAS)", trigger: "avviso", filtroAvviso: "ROAS", soloCritici: "no",
    taskTitolo: "ROAS sotto l'obiettivo della campagna", taskPriorita: "alta", taskGiorni: 2,
    perche: "Quarta soglia per le campagne a vendita: il ritorno è sotto il minimo impostato sulla campagna." },
  { nome: "ADA · Costo per lead oltre obiettivo (CPL)", trigger: "avviso", filtroAvviso: "CPL", soloCritici: "no",
    taskTitolo: "CPL sopra l'obiettivo della campagna", taskPriorita: "alta", taskGiorni: 2,
    perche: "Quarta soglia nelle campagne a contatti: il lead è il primo punto di conversione e costa più del tetto. Qui si interviene prima che il problema arrivi al CPA." },
  { nome: "ADA · Costo per risultato oltre obiettivo (CPA)", trigger: "avviso", filtroAvviso: "CPA", soloCritici: "no",
    taskTitolo: "CPA sopra l'obiettivo della campagna", taskPriorita: "alta", taskGiorni: 2,
    perche: "Quarta soglia per le campagne a contatto: il costo per risultato ha superato il massimo impostato." },
  { nome: "ADA · Peggioramento costante (tendenza)", trigger: "avviso", filtroAvviso: "Tendenza", soloCritici: "no",
    taskTitolo: "Tre rilevazioni consecutive in peggioramento", taskPriorita: "alta", taskGiorni: 2,
    perche: "Il test di tendenza: non conta il valore di oggi ma la direzione. Serve ad agire prima che la soglia venga superata." },
  { nome: "ADA · Pubblico saturo (frequenza)", trigger: "avviso", filtroAvviso: "Frequenza", soloCritici: "no",
    taskTitolo: "Frequenza oltre soglia: valuta nuovo pubblico o creativo", taskPriorita: "media", taskGiorni: 3,
    perche: "Quinta soglia: la frequenza da sola non è un problema, lo diventa quando si accompagna a risultati in calo." },
  { nome: "ADA · Erogazione anomala", trigger: "avviso", filtroAvviso: "Erogazione", soloCritici: "no",
    taskTitolo: "Il budget non viene speso come previsto", taskPriorita: "alta", taskGiorni: 2,
    perche: "Sesta soglia: una campagna che non eroga non si giudica sul rendimento, e non si scala finché non consuma quello che ha." },
  { nome: "ADA · Avviso critico su qualunque metrica", trigger: "avviso", filtroAvviso: "", soloCritici: "si",
    taskTitolo: "Avviso critico da esaminare oggi", taskPriorita: "urgente", taskGiorni: 1,
    perche: "Rete di sicurezza: qualunque avviso critico, anche su metriche non coperte dalle regole sopra, diventa un task urgente." },
];

const TRIGGER = [
  { v: "avviso", l: "Si accende un avviso su una campagna" },
  { v: "metrica", l: "Una metrica supera una soglia" },
  { v: "budget", l: "Il budget erogato supera una percentuale" },
  { v: "task", l: "Un task va in ritardo" },
  { v: "chiusura", l: "Viene registrata una chiusura" },
  { v: "periodo", l: "Il mese non ha ancora dati commerciali" },
];
const OPERATORI = [{ v: "gt", l: "è maggiore di" }, { v: "lt", l: "è minore di" }];
const METRICHE_TRIGGER = [
  { v: "ctr", l: "CTR (%)" }, { v: "cpc", l: "CPC" }, { v: "cpm", l: "CPM" }, { v: "cpl", l: "CPL" }, { v: "cpa", l: "CPA" },
  { v: "roas", l: "ROAS (×)" }, { v: "freq", l: "Frequenza" }, { v: "sp", l: "Spesa" }, { v: "cv", l: "Conversioni" },
];
const FILTRI_AVVISO = [
  { v: "", l: "Qualsiasi avviso" }, { v: "ROAS", l: "Solo ROAS" }, { v: "CPL", l: "Solo CPL" }, { v: "CPA", l: "Solo CPA" }, { v: "CTR", l: "Solo CTR" },
  { v: "Frequenza", l: "Solo frequenza" }, { v: "CPM", l: "Solo CPM" }, { v: "Conversioni", l: "Solo conversioni" },
  { v: "Erogazione", l: "Solo erogazione" },
];
const SN = [{ v: "si", l: "Sì" }, { v: "no", l: "No" }];
const TIPI_CREATIVO = [{ v: "adset", l: "Ad set" }, { v: "annuncio", l: "Annuncio" }];
const TIPI_NOTA = [
  { v: "budget", l: "Budget" }, { v: "creativita", l: "Creatività" }, { v: "pubblico", l: "Pubblico" },
  { v: "offerta", l: "Offerta o prezzo" }, { v: "tracciamento", l: "Tracciamento" },
  { v: "esterno", l: "Fattore esterno" }, { v: "altro", l: "Altro" },
];
const STATI_ESPERIMENTO = [
  { v: "corso", l: "In corso" }, { v: "vinto", l: "Confermato" },
  { v: "perso", l: "Smentito" }, { v: "nullo", l: "Inconcludente" },
];

const TERMINI_PAGAMENTO = [
  { v: "ricezione", l: "Alla ricezione" }, { v: "15", l: "15 giorni" }, { v: "30", l: "30 giorni" },
  { v: "30fm", l: "30 giorni fine mese" }, { v: "60", l: "60 giorni" }, { v: "60fm", l: "60 giorni fine mese" },
  { v: "anticipato", l: "Anticipato" },
];
const METODI_PAGAMENTO = [
  { v: "bonifico", l: "Bonifico" }, { v: "sdd", l: "Addebito SDD" }, { v: "carta", l: "Carta" }, { v: "altro", l: "Altro" },
];

/* Le cinque fasi del processo d'acquisto: servono a capire se il messaggio
   copre tutto il percorso o salta direttamente dal problema all'acquisto. */
const STADI = [
  { v: "1", l: "1 · Riconoscimento del problema", b: "Problema" },
  { v: "2", l: "2 · Raccolta di informazioni", b: "Informazioni" },
  { v: "3", l: "3 · Valutazione delle soluzioni", b: "Valutazione" },
  { v: "4", l: "4 · Acquisto", b: "Acquisto" },
  { v: "5", l: "5 · Dopo l'acquisto", b: "Post-acquisto" },
];

const PASSI_AOA = [
  { k: "s1", l: "Piano di messaggio completo", d: "Messaggi per le fasi 1, 2, 3 e 4 coordinati tra annunci, funnel, follow-up e vendita. È lavoro di pianificazione: meglio è fatto, più alta è la probabilità che le campagne funzionino appena pubblicate." },
  { k: "s2", l: "Piano di conversione scelto", d: "Funnel, follow-up, sistema di vendita e strategia pubblicitaria decisi insieme, non uno alla volta." },
  { k: "s3", l: "Asset creati", d: "Funnel, sequenze di follow-up, annunci e pubblici pronti, passati da revisione." },
  { k: "s4", l: "Dati attendibili", d: "Tracciamento installato in modo pulito e validato contro i numeri reali dell'azienda. L'obiettivo non è la perfezione ma conoscere lo scarto: una precisione fra l'85 e il 95% è il livello a cui si può comprare media." },
  { k: "s5", l: "Test di validazione lanciato", d: "Campagna di test iniziale, più le campagne di metà e fondo imbuto quando servono. Qui si perde denaro tagliando gli angoli." },
  { k: "s6", l: "Operatività e scala", d: "Regole chiare su quanto spendere, cosa è scalabile e quando spegnere: è il framework a sei soglie." },
  { k: "s7", l: "Ottimizzazione del piano", d: "Trovare il punto di fallimento del percorso di conversione e lavorare solo su quello." },
];

const CONSAPEVOLEZZA = [
  { v: "1", l: "1 · Conosce già la soluzione e la cerca" },
  { v: "2", l: "2 · Conosce il prodotto ma non è convinto" },
  { v: "3", l: "3 · Sa di avere il problema, non conosce le soluzioni" },
  { v: "4", l: "4 · Avverte un sintomo, non ha nominato il problema" },
  { v: "5", l: "5 · Del tutto ignaro" },
];
const GANCI = {
  "1": "Dichiara direttamente il risultato che il mercato desidera, con il numero o la condizione precisa.",
  "2": "Fai la promessa più grande che puoi sostenere, alzando l'asticella rispetto a chi già conosce la categoria.",
  "3": "Presenta un meccanismo unico: non il risultato, ma il modo nuovo con cui lo ottieni.",
  "4": "Presenta un meccanismo migliorato o aggiornato rispetto a quello che il mercato ha già visto fallire.",
  "5": "Serve una strategia nuova: parti dal sintomo quotidiano, non dal prodotto né dal problema nominato.",
};

const CATEGORIE_FILE = [
  { v: "strategia", l: "Strategia" }, { v: "procedura", l: "Procedura" }, { v: "brief", l: "Brief" },
  { v: "creativita", l: "Creatività" }, { v: "contratto", l: "Contratto" }, { v: "report", l: "Report" },
  { v: "dati", l: "Dati e fogli" }, { v: "altro", l: "Altro" },
];
const TIPI_FILE = [{ v: "link", l: "Collegamento esterno" }, { v: "documento", l: "Documento scritto qui" }];

/* Modulo giornaliero per i venditori: le colonne sono anche le chiavi di lettura */
const COL_DIARIO = [
  { k: "data", l: "Data", ex: "" },
  { k: "venditore", l: "Venditore", ex: "Mario Rossi" },
  { k: "funnel", l: "Funnel di provenienza", ex: "VSL" },
  { k: "lead", l: "Lead ricevuti", ex: "8" },
  { k: "chiamate", l: "Chiamate fatte", ex: "22" },
  { k: "discoveryFissate", l: "Discovery fissate", ex: "4" },
  { k: "discoverySvolte", l: "Discovery svolte", ex: "3" },
  { k: "demoFissate", l: "Demo fissate", ex: "2" },
  { k: "demoSvolte", l: "Demo svolte", ex: "1" },
  { k: "offerte", l: "Offerte presentate", ex: "1" },
  { k: "chiusure", l: "Chiusure", ex: "1" },
  { k: "valoreContratto", l: "Valore contratto", ex: "4500" },
  { k: "incassato", l: "Incassato", ex: "1500" },
  { k: "oreContatto", l: "Ore al primo contatto", ex: "3" },
  { k: "note", l: "Note", ex: "due no show, richiamo domani" },
];
const NUM_DIARIO = ["lead", "chiamate", "discoveryFissate", "discoverySvolte", "demoFissate", "demoSvolte", "offerte", "chiusure", "valoreContratto", "incassato", "oreContatto"];

/* --- funnel --- */
const TIPI_PASSO = [
  { v: "annuncio", l: "Annuncio" }, { v: "pagina", l: "Landing page" }, { v: "vsl", l: "Video di vendita" },
  { v: "form", l: "Modulo" }, { v: "calendario", l: "Prenotazione" }, { v: "chiamata", l: "Chiamata o demo" },
  { v: "checkout", l: "Checkout" }, { v: "upsell", l: "Upsell" }, { v: "email", l: "Email" },
];
const OBIETTIVI_FUNNEL = [
  { v: "appuntamento", l: "Prenotare chiamate" }, { v: "lead", l: "Raccogliere lead" }, { v: "vendita", l: "Vendere online" },
];
const DIAGNOSI_PASSO = {
  annuncio: { causa: "Creatività consumata o pubblico poco pertinente: chi vede l'annuncio non si riconosce nel problema.", azione: "Cambia il gancio dei primi tre secondi e verifica che la promessa dell'annuncio sia la stessa della pagina che segue." },
  pagina: { causa: "Quasi sempre è disallineamento: la pagina promette qualcosa di diverso dall'annuncio, o mette il prezzo prima del valore.", azione: "Riprendi nel titolo le stesse parole dell'annuncio, sposta la prova sociale sopra la piega e misura la velocità di caricamento da mobile." },
  vsl: { causa: "Il video perde le persone nei primi trenta secondi, oppure è troppo lungo per la temperatura del traffico.", azione: "Guarda dove cade la curva di visualizzazione e riscrivi quel punto. Su traffico freddo, accorcia e anticipa il beneficio." },
  form: { causa: "Troppi campi o richieste premature: ogni campo in più costa conversioni, e i dati sensibili vanno chiesti dopo.", azione: "Riduci ai campi che servono davvero per la prima chiamata e spiega in una riga cosa succede dopo l'invio." },
  calendario: { causa: "Slot scomodi, troppa attesa tra oggi e la prima disponibilità, oppure passaggi extra prima di vedere il calendario.", azione: "Apri slot più vicini nel tempo, mostra il calendario direttamente in pagina e manda subito conferma e promemoria." },
  chiamata: { causa: "Chi prenota non è qualificato oppure arriva senza aver capito cosa lo aspetta: è il classico buco dello show up.", azione: "Aggiungi qualifica prima della prenotazione, promemoria a 24 e a 1 ora, e un messaggio che ricorda perché ha prenotato." },
  checkout: { causa: "Costi che compaiono alla fine, pochi metodi di pagamento, o un modulo troppo lungo.", azione: "Mostra tutti i costi prima, offri più metodi di pagamento e riduci il checkout a una schermata sola." },
  upsell: { causa: "L'offerta non è collegata a quello appena comprato, o arriva quando l'attenzione è già finita.", azione: "Proponi un complemento naturale dell'acquisto, con una sola scelta chiara e un prezzo inferiore all'ordine principale." },
  email: { causa: "Oggetto debole, invio troppo distante dall'iscrizione o contenuto che non riprende la promessa iniziale.", azione: "Manda la prima email entro pochi minuti, riprendi la promessa nell'oggetto e metti un solo invito all'azione." },
};
const MODELLI_FUNNEL = [
  { v: "", l: "— parti da zero —", passi: [] },
  { v: "call", l: "Traffico → VSL → prenotazione → chiamata", passi: [["Annuncio", "annuncio"], ["Landing con VSL", "vsl"], ["Modulo di qualifica", "form"], ["Calendario", "calendario"], ["Chiamata svolta", "chiamata"]] },
  { v: "lead", l: "Traffico → lead magnet → email → chiamata", passi: [["Annuncio", "annuncio"], ["Pagina lead magnet", "pagina"], ["Modulo", "form"], ["Sequenza email", "email"], ["Chiamata svolta", "chiamata"]] },
  { v: "webinar", l: "Traffico → registrazione → webinar → offerta", passi: [["Annuncio", "annuncio"], ["Pagina di registrazione", "pagina"], ["Presenza al webinar", "vsl"], ["Pagina offerta", "pagina"], ["Checkout", "checkout"]] },
  { v: "shop", l: "Traffico → prodotto → checkout → upsell", passi: [["Annuncio", "annuncio"], ["Scheda prodotto", "pagina"], ["Checkout", "checkout"], ["Upsell", "upsell"]] },
];
const mese = (d) => d.slice(0, 7);
const meseIt = (s) => (s ? new Date(s + "-01T00:00:00").toLocaleDateString(LOCALE, { month: "short", year: "2-digit" }) : "—");

const uid = () => Math.random().toString(36).slice(2, 9);
/* Date sempre in ora locale: toISOString() lavora in UTC e in Italia sposterebbe il giorno */
const isoLocale = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const oggi = () => isoLocale(new Date());
const sommaGiorni = (s, g) => { const d = new Date(s + "T00:00:00Z"); d.setUTCDate(d.getUTCDate() + g); return d.toISOString().slice(0, 10); };
const piu = (g) => sommaGiorni(oggi(), g);

/* Contenuti pronti: usati nei dati iniziali e recuperabili con un pulsante */
const PROCEDURE_BASE = () => [
  { id: uid(), clienteId: "", titolo: "Procedura: rilevazione settimanale", categoria: "procedura", tipo: "documento", url: "", tag: "routine", aggiornato: oggi(),
    testo: "Ogni lunedì mattina, prima di qualsiasi ottimizzazione:\n\n1. Apri ogni piattaforma ed esporta il report campagne del periodo corrente.\n2. Importa dati → scegli il profilo salvato → carica il file.\n3. Advertising → Salva rilevazione. Senza questo passaggio la colonna Andamento resta vuota e i due avvisi sul peggioramento non scattano mai.\n4. Guarda gli avvisi: quelli critici o si risolvono subito o diventano un task con una data.\n5. Scrivi nel Diario ogni modifica fatta, anche piccola.\n\nTempo previsto: venti minuti per tre account." },
  { id: uid(), clienteId: "", titolo: "Procedura: chiusura del mese", categoria: "procedura", tipo: "documento", url: "", tag: "routine", aggiornato: oggi(),
    testo: "Entro il quinto giorno del mese nuovo:\n\n1. Ciclo di vendita: crea il periodo del mese appena chiuso, con lead, chiamate, discovery e demo (fissate e svolte), slot disponibili e ore medie al primo contatto.\n2. Chiusure: inserisci una scheda per ogni cliente chiuso, con ad set e annuncio scritti esattamente come in piattaforma.\n3. Report: genera e manda al cliente.\n4. Economia: controlla margine per cliente e concentrazione del fatturato.\n5. Backup: scarica il file e mettilo nella cartella sincronizzata." },
  { id: uid(), clienteId: "", titolo: "Procedura: nuovo cliente, primi sette giorni", categoria: "procedura", tipo: "documento", url: "", tag: "onboarding", aggiornato: oggi(),
    testo: "Prima di spendere un euro:\n\n1. Scheda cliente completa, margine lordo e resi compresi: senza quelli il ROAS non dice se il cliente guadagna.\n2. Verifica del tracciamento con una conversione di test, prima di alzare i budget.\n3. Benchmark concordati con il cliente e scritti nelle campagne, non tenuti a mente.\n4. Capacità commerciale: quante discovery riesce a gestire in un mese? È il tetto vero del budget.\n5. Collegamento alla cartella Drive del cliente qui in File, e strategia allegata.\n6. Prima annotazione nel Diario: da dove partiamo e cosa ci aspettiamo." },
  { id: uid(), clienteId: "", titolo: "Framework delle decisioni sugli annunci (Jason Hornung)", categoria: "strategia", tipo: "documento", url: "", tag: "meta, benchmark", aggiornato: oggi(),
    testo: "Sintesi del framework a sei soglie di Jason Hornung (Academy of Advertising), pensato per account Meta con spesa mensile importante. Le soglie vanno lette in ordine: la prima è un cancello, le altre hanno senso solo dopo.\n\n1. SOGLIA DI SPESA PRIMA DI GIUDICARE\nUn annuncio deve aver speso due volte il valore medio dell'ordine (o due volte il costo per lead obiettivo) prima di qualsiasi valutazione. Sotto quella cifra l'algoritmo non ha segnale sufficiente: qualunque decisione è una sensazione, non una lettura. Azione: annota la tendenza, non agire, torna quando la soglia è raggiunta.\n\n2. CTR — FALLIMENTO DEL GANCIO\nCTR sul link sotto lo 0,8% dopo la soglia. È un problema di creatività, non di pubblico. Azione: spegni la creatività, tieni il pubblico, riscrivi il gancio con un angolo diverso e la stessa offerta.\n\n3. CPM — SALUTE DELL'ASTA\nCPM oltre il 40% sopra la media dell'account, dopo la soglia. L'algoritmo paga un sovrapprezzo perché non riceve abbastanza segnali positivi. Azione: sostituisci la creatività dentro lo stesso ad set, il pubblico può essere ancora valido.\n\n4. CANCELLO ROAS / CPL\nSotto il ROAS obiettivo (o sopra il CPL obiettivo) dopo la soglia, senza tendenza in miglioramento negli ultimi tre giorni. La clausola sui tre giorni evita di spegnere annunci disturbati temporaneamente dalla consegna. Azione: spegni e documenta l'angolo fallito; non ritestare lo stesso gancio sullo stesso pubblico.\n\n5. SLOPE TEST — DECLINO CONTROLLATO\nL'annuncio era a obiettivo al lancio ma ROAS o CPL peggiorano da tre o più settimane consecutive senza cause esterne. È affaticamento al rallentatore: l'algoritmo continua a consegnare per via dello storico mentre l'interesse reale crolla. Aspettare il fondo costa spesa ogni settimana. Azione: limita il budget e lancia subito una creatività sfidante.\n\n6. L'ECCEZIONE — FREQUENZA ALTA MA ANCORA PROFITTEVOLE\nFrequenza sopra 4,5 con risultati ancora a obiettivo o superiori. Significa che la creatività ha una tenuta insolita. Quasi tutti spengono per abitudine a 3 o 4: se i numeri dicono il contrario, hanno ragione i numeri.\n\nNOTA SU COME È APPLICATO QUI\nQueste soglie sono già dentro il motore di avvisi: la prima fa da cancello e sospende i giudizi finché la spesa non è sufficiente, la sesta impedisce l'allarme sulla frequenza quando i risultati reggono. Il valore medio dell'ordine si imposta nella scheda cliente, il CPA obiettivo nei benchmark della campagna.\n\nFonte: The Ad Decision Framework, Academy of Advertising — academyofadvertising.ai" },
  { id: uid(), clienteId: "", titolo: "I quattro annunci e il piano di messaggio (AOA — Jason Hornung)", categoria: "strategia", tipo: "documento", url: "", tag: "messaggio, creativita", aggiornato: oggi(),
    testo: "Sintesi della parte sul messaggio del metodo AOA. Il principio: il processo d'acquisto ha cinque fasi (riconoscimento del problema, raccolta di informazioni, valutazione delle soluzioni, acquisto, dopo l'acquisto) e la maggior parte degli account porta le persone dalla fase 1 alla fase 4 con un annuncio solo. Questo messaggio corto produce tre sintomi riconoscibili: campagne che si rompono quando il budget sale, qualità dei lead bassa, e appuntamenti presi da chi poi non si presenta.\n\nLA STRUTTURA DEI QUATTRO ANNUNCI\nUn annuncio per ciascuna delle fasi 1-4. Per la fase 5 non si fanno annunci.\n\nFASE 1 — traffico freddo, problema e soluzione\nCorpo: affermazione o domanda che identifica il dolore o il guadagno, poi la promessa di risolverlo, poi tre grandi benefici in elenco, poi l'invito all'azione.\nCreativita: titolo su blocco di colore.\nTitolo: dedicato esplicitamente al pubblico, del tipo Solo per [mercato].\n\nFASE 2 — traffico tiepido, informazione utile\nCorpo: Leggi questo prima di acquistare [cosa vendi].\nCreativita: formato leggero, tono da contenuto e non da pubblicita.\nTitolo: la guida definitiva all'acquisto di [cosa vendi].\n\nFASE 3 — traffico tiepido, chi siamo\nCorpo: missione dichiarata al mercato, superamento dell'obiezione piu grande, affermazione del beneficio principale.\nCreativita: video dimostrativo del prodotto o servizio.\nTitolo: Perche scegliere [azienda]?\n\nFASE 4 — traffico caldo, offerta\nCorpo: proiezione nel futuro dello stato emotivo dopo aver ottenuto il risultato.\nCreativita: immagine del prodotto.\nTitolo: l'offerta piu un invito di una parola.\n\nLE DOMANDE CHE VENGONO PRIMA\nChi e il mercato. Quale dolore risolvi o quale guadagno dai. Qual e la tua soluzione unica. I tre grandi benefici. Le tre grandi obiezioni. Come sara il futuro del cliente dopo il risultato. Cosa vendi. La missione. L'offerta.\n\nDUE DISTINZIONI CHE VALGONO PIU DI TUTTO IL RESTO\nDolori di chi compra contro dolori dei curiosi: entrambi veri, ma solo i primi sono attuali e urgenti. Costruire il messaggio sui secondi riempie il calendario di persone che non comprano.\nConsapevolezza del mercato: il gancio cambia a seconda di quanto il mercato conosce gia il problema e le soluzioni. A un mercato che conosce la soluzione si dichiara il risultato; a uno che ha gia visto fallire i meccanismi noti serve un meccanismo nuovo o aggiornato; a uno del tutto ignaro si parte dal sintomo quotidiano.\n\nCOME E APPLICATO QUI\nLa sezione Messaggio di ogni cliente raccoglie queste risposte e compone la frase che tiene insieme tutto, poi genera i quattro annunci gia scritti con le tue parole e li puo creare in Advertising, uno per fase. Il collegamento con il commerciale e automatico: quando lo show up alle demo e basso e la spesa sulle fasi 2 e 3 e vicina a zero, l'app te lo dice.\n\nFonte: AOA Core 4 Ad Strategy Guide e Install Guide Step 1 Worksheets, Academy of Advertising." },
  { id: uid(), clienteId: "", titolo: "Il metodo in sette passi (AOA Way — Jason Hornung)", categoria: "strategia", tipo: "documento", url: "", tag: "metodo, impianto", aggiornato: oggi(),
    testo: "Sintesi del metodo AOA Way di Jason Hornung: cinque passi prima di pubblicare, due dopo. Il framework a sei soglie che trovi negli avvisi è il sesto passo; da solo non basta, perché dice quando spegnere ma non cosa costruire.\n\nPRIMA DI PUBBLICARE\n\n1. PIANO DI MESSAGGIO COMPLETO\nIl processo d'acquisto ha cinque fasi: riconoscimento del problema, raccolta di informazioni, valutazione delle soluzioni, acquisto, dopo l'acquisto. La maggior parte degli account porta le persone dalla fase 1 alla fase 4 con un solo annuncio, saltando la 2 e la 3. È il messaggio corto, ed è la causa più comune del logoramento rapido delle campagne. Il piano completo prevede messaggi per le fasi 1, 2, 3 e 4, coordinati tra annunci, funnel, follow-up e sistema di vendita.\n\n2. PIANO DI CONVERSIONE\nFunnel, follow-up, vendita e strategia pubblicitaria scelti insieme come un sistema unico, non assemblati un pezzo alla volta.\n\n3. ASSET\nFunnel, sequenze di follow-up, annunci e pubblici creati e passati da revisione prima del lancio.\n\n4. DATI ATTENDIBILI\nTre principi: non affidarsi a software di tracciamento di terze parti; nessun sistema darà mai dati accurati al 100% e sistemi diversi non si riconcilieranno mai del tutto; i dati vanno letti in modo complessivo, guardando l'effetto sul risultato reale dell'azienda. L'obiettivo è conoscere lo scarto e validarlo con regolarità contro i numeri veri. Un'installazione pulita del tracciamento porta l'accuratezza dei gestori pubblicitari fra l'85 e il 95%: è il livello a cui si può comprare media con fiducia.\n\n5. TEST DI VALIDAZIONE\nSi lancia una campagna di test iniziale, più le campagne di metà e fondo imbuto se previste dal piano. È il punto in cui l'entusiasmo fa tagliare gli angoli, ed è il taglio più costoso. Meta è indicata come piattaforma di test più affidabile ed economica.\n\nDOPO LA PUBBLICAZIONE\n\n6. OPERARE, SCALARE, MANTENERE\nRegole chiare su quanto spendere, come capire se un annuncio è scalabile e quando spegnerlo. Qui vivono le sei soglie: spesa minima prima di giudicare, CTR, CPM contro la media dell'account, cancello ROAS/CPL con la clausola sulla tendenza, slope test, eccezione sulla frequenza.\n\n7. OTTIMIZZARE IL PIANO DI CONVERSIONE\nSi individua il punto di fallimento del percorso e si lavora solo su quello, invece di ritoccare tutto insieme. La sezione Funnel di questa app fa esattamente questo calcolo.\n\nCOME È APPLICATO QUI\nLa fase del messaggio si assegna a ogni ad set e annuncio: in Advertising trovi la spesa divisa per fase e un allarme se stai finanziando solo problema e acquisto. Lo scarto fra conversioni dichiarate dalle piattaforme e lead reali si controlla nel Ciclo di vendita. I sette passi sono una lista di installazione nella scheda di ogni cliente.\n\nFonte: The AOA Way Install Guide, Academy of Advertising." },
];

const REGOLE_BASE = () => [
  { id: uid(), nome: "Frequenza fuori soglia → task creatività", attiva: "si", clienteId: "", trigger: "avviso", filtroAvviso: "Frequenza", soloCritici: "no", metrica: "freq", operatore: "gt", valore: 3.5, giorno: 5, azTask: "si", taskTitolo: "Rinnovare le creatività", taskArea: "ar2", taskPriorita: "alta", taskGiorni: 2, azWebhook: "no", url: "" },
  { id: uid(), nome: "Budget quasi esaurito → avvisami", attiva: "si", clienteId: "", trigger: "budget", filtroAvviso: "", soloCritici: "no", metrica: "sp", operatore: "gt", valore: 85, giorno: 5, azTask: "si", taskTitolo: "Rivedere il budget giornaliero", taskArea: "ar1", taskPriorita: "urgente", taskGiorni: 1, azWebhook: "no", url: "" },
  { id: uid(), nome: "Nuova chiusura → notifica esterna", attiva: "no", clienteId: "", trigger: "chiusura", filtroAvviso: "", soloCritici: "no", metrica: "sp", operatore: "gt", valore: 0, giorno: 5, azTask: "no", taskTitolo: "", taskArea: "ar5", taskPriorita: "media", taskGiorni: 3, azWebhook: "si", url: "" },
  { id: uid(), nome: "Mese senza dati commerciali → promemoria", attiva: "si", clienteId: "", trigger: "periodo", filtroAvviso: "", soloCritici: "no", metrica: "sp", operatore: "gt", valore: 0, giorno: 8, azTask: "si", taskTitolo: "Chiedere i numeri commerciali del mese", taskArea: "ar3", taskPriorita: "media", taskGiorni: 2, azWebhook: "no", url: "" },
];

const SEED = () => ({
  agenzia: { nome: "NC", sottotitolo: "Media buying", logo: "", costoOrario: 60 },
  aree: AREE_DEFAULT,
  clienti: [
    { id: "z1", nome: "Cantine del Chienti", settore: "Vino / e-commerce", citta: "Tolentino", stato: "attivo", fee: 1800, budgetMese: 12000, margineLordo: 42, resi: 6, aov: 96, logo: "", note: "Stagionalità forte da ottobre. Margine alto sulle bottiglie premium.", creato: piu(-210),
      ragioneSociale: "Cantine del Chienti S.r.l.", piva: "01234567890", codiceFiscale: "01234567890", indirizzo: "Contrada Rambona 14", cap: "62029", cittaFatt: "Tolentino", provincia: "MC", paese: "Italia", sdi: "M5UXCR1", pec: "amministrazione@pec.cantinechienti.it", emailAmm: "amministrazione@cantinechienti.it", referenteAmm: "Luca Petrini", termini: "30fm", metodo: "bonifico" },
    { id: "z2", nome: "Tecnoplast Marche", settore: "B2B manifatturiero", citta: "Fermo", stato: "attivo", fee: 2200, budgetMese: 6000, margineLordo: 35, resi: 0, aov: 0, logo: "", note: "Solo lead qualificati, il commerciale richiama entro 24h.", creato: piu(-120),
      ragioneSociale: "Tecnoplast Marche S.p.A.", piva: "09876543210", codiceFiscale: "09876543210", indirizzo: "Via dell'Industria 8", cap: "63900", cittaFatt: "Fermo", provincia: "FM", paese: "Italia", sdi: "", pec: "tecnoplast@pec.it", emailAmm: "contabilita@tecnoplast.it", referenteAmm: "Silvia Marchetti", termini: "60", metodo: "bonifico" },
    { id: "z3", nome: "Studio Bellini", settore: "Servizi professionali", citta: "Macerata", stato: "onboarding", fee: 900, budgetMese: 1500, margineLordo: 70, resi: 0, aov: 0, logo: "", note: "Tracciamento ancora da sistemare.", creato: piu(-18),
      ragioneSociale: "Studio Bellini S.t.p.", piva: "", codiceFiscale: "", indirizzo: "", cap: "", cittaFatt: "", provincia: "", paese: "Italia", sdi: "", pec: "", emailAmm: "", referenteAmm: "", termini: "30", metodo: "bonifico" },
  ],
  contatti: [
    { id: "c1", clienteId: "z1", tipo: "referente", nome: "Sara Lupini", azienda: "Cantine del Chienti", ruolo: "Marketing manager", email: "sara@cantinechienti.it", telefono: "0733 210044", stato: "cliente", campagnaId: "", note: "Approva creatività e budget." },
    { id: "c2", clienteId: "z1", tipo: "lead", nome: "Enoteca Bertelli", azienda: "Bertelli srl", ruolo: "Titolare", email: "info@bertelli.it", telefono: "071 445566", stato: "qualificato", campagnaId: "k1", note: "Interessata alla fornitura ristorazione." },
    { id: "c3", clienteId: "z2", tipo: "referente", nome: "Paolo Renzi", azienda: "Tecnoplast Marche", ruolo: "Titolare", email: "p.renzi@tecnoplast.it", telefono: "0734 556677", stato: "cliente", campagnaId: "", note: "" },
    { id: "c4", clienteId: "z2", tipo: "lead", nome: "Andrea Sacchi", azienda: "Officine Sacchi", ruolo: "Direttore tecnico", email: "a.sacchi@offsacchi.it", telefono: "0721 334455", stato: "opportunita", campagnaId: "k3", note: "Ha chiesto preventivo impianto." },
    { id: "c5", clienteId: "z2", tipo: "lead", nome: "Marta Coppola", azienda: "Adriatica Logistica", ruolo: "Resp. acquisti", email: "m.coppola@adrlog.it", telefono: "071 998412", stato: "contattato", campagnaId: "k4", note: "" },
    { id: "c6", clienteId: "z3", tipo: "referente", nome: "Anna Bellini", azienda: "Studio Bellini", ruolo: "Socia", email: "a.bellini@studiobellini.it", telefono: "0733 887711", stato: "cliente", campagnaId: "", note: "" },
  ],
  trattative: [
    { id: "t1", clienteId: "z1", titolo: "Fornitura ristorazione", contattoId: "c2", valore: 9800, fase: "proposta", chiusura: piu(20), campagnaId: "k1", note: "" },
    { id: "t2", clienteId: "z2", titolo: "Impianto pilota", contattoId: "c4", valore: 31000, fase: "negoziazione", chiusura: piu(25), campagnaId: "k3", note: "Budget confermato." },
    { id: "t3", clienteId: "z2", titolo: "Fornitura ricambi", contattoId: "c5", valore: 7400, fase: "qualificato", chiusura: piu(50), campagnaId: "k4", note: "" },
    { id: "t4", clienteId: "z1", titolo: "Ordine natalizio corporate", contattoId: "c2", valore: 5200, fase: "vinto", chiusura: piu(-20), campagnaId: "k1", note: "" },
  ],
  campagne: [
    { id: "k1", clienteId: "z1", nome: "Catalogo autunno — vendite", piattaforma: "meta", obiettivo: "vendite", stato: "attiva", inizio: piu(-24), fine: piu(6), budgetGiorno: 265, budget: 8000, spesa: 6412.4, impression: 812450, click: 6903, clickLink: 5120, lead: 0, conversioni: 118, ricavi: 14238.5, frequenza: 4.32, bmk: { roasMin: 3, cpaMax: 0, ctrMin: 1, freqMax: 3.5, cpmMax: 0 } },
    { id: "k2", clienteId: "z1", nome: "Brand search", piattaforma: "google", obiettivo: "vendite", stato: "attiva", inizio: piu(-60), fine: piu(30), budgetGiorno: 28, budget: 2500, spesa: 1183.9, impression: 41200, click: 3105, clickLink: 2980, lead: 0, conversioni: 96, ricavi: 9142.2, frequenza: 1.24, bmk: { roasMin: 4, cpaMax: 0, ctrMin: 4, freqMax: 0, cpmMax: 0 } },
    { id: "k3", clienteId: "z2", nome: "Lead impianti — pubblico freddo", piattaforma: "linkedin", obiettivo: "lead", stato: "attiva", funnelId: "f1", inizio: piu(-30), fine: piu(15), budgetGiorno: 100, budget: 4500, spesa: 3604.8, impression: 240300, click: 1452, clickLink: 1030, lead: 21, conversioni: 21, ricavi: 0, frequenza: 2.41, bmk: { roasMin: 0, cpaMax: 120, ctrMin: 0.5, freqMax: 3, cpmMax: 0 } },
    { id: "k4", clienteId: "z2", nome: "Retargeting sito", piattaforma: "meta", obiettivo: "lead", stato: "attiva", funnelId: "f1", inizio: piu(-30), fine: piu(15), budgetGiorno: 27, budget: 1200, spesa: 981.3, impression: 96400, click: 1902, clickLink: 1240, lead: 14, conversioni: 14, ricavi: 0, frequenza: 3.12, bmk: { roasMin: 0, cpaMax: 60, ctrMin: 1.2, freqMax: 3, cpmMax: 0 } },
    { id: "k5", clienteId: "z3", nome: "Primo test locale", piattaforma: "meta", obiettivo: "lead", stato: "test", inizio: piu(-12), fine: piu(18), budgetGiorno: 50, budget: 1500, spesa: 431.7, impression: 88200, click: 619, clickLink: 402, lead: 0, conversioni: 0, ricavi: 0, frequenza: 1.63, bmk: { roasMin: 0, cpaMax: 40, ctrMin: 1, freqMax: 3, cpmMax: 0 } },
  ],
  task: [
    { id: "s1", clienteId: "z1", areaId: "ar2", titolo: "Tre nuove creatività catalogo autunno", note: "La frequenza è a 4,32: servono ganci diversi, non varianti.", stato: "incorso", priorita: "urgente", inizio: piu(-3), scadenza: piu(-1), stima: 4, ricorrenza: "no", checklist: [{ id: "q1", t: "Brief copy", fatto: true }, { id: "q2", t: "Montaggio video 9:16", fatto: false }, { id: "q3", t: "Caricamento e naming", fatto: false }], creato: piu(-5) },
    { id: "s2", clienteId: "z2", areaId: "ar3", titolo: "Report mensile e call di allineamento", note: "", stato: "dafare", priorita: "alta", inizio: oggi(), scadenza: oggi(), stima: 2, ricorrenza: "30", checklist: [], creato: piu(-30) },
    { id: "s3", clienteId: "z3", areaId: "ar4", titolo: "Verificare pixel e conversioni", note: "Acquisto di test prima di alzare il budget.", stato: "attesa", priorita: "alta", inizio: piu(-2), scadenza: piu(2), stima: 1.5, ricorrenza: "no", checklist: [], creato: piu(-4) },
    { id: "s4", clienteId: "z2", areaId: "ar5", titolo: "Richiamare Officine Sacchi", note: "", stato: "dafare", priorita: "media", inizio: oggi(), scadenza: piu(1), stima: 0.5, ricorrenza: "no", checklist: [], creato: piu(-1) },
    { id: "s5", clienteId: "", areaId: "ar6", titolo: "Fatture del mese", note: "", stato: "dafare", priorita: "bassa", inizio: piu(4), scadenza: piu(8), stima: 1, ricorrenza: "30", checklist: [], creato: piu(-2) },
    { id: "s6", clienteId: "z1", areaId: "ar1", titolo: "Ribilanciare budget verso Brand search", note: "", stato: "dafare", priorita: "alta", inizio: oggi(), scadenza: piu(2), stima: 1, ricorrenza: "no", checklist: [], creato: oggi() },
  ],
  vendite: [
    { id: "v1", clienteId: "z2", periodo: mese(piu(-62)), lead: 96, chiamate: 210, discoveryFissate: 38, discoverySvolte: 27, demoFissate: 19, demoSvolte: 13, offerte: 11, chiusure: 4, valoreContratto: 46000, incassato: 28000, spesa: 3900, slot: 50, oreContatto: 6 },
    { id: "v2", clienteId: "z2", periodo: mese(piu(-31)), lead: 112, chiamate: 244, discoveryFissate: 45, discoverySvolte: 29, demoFissate: 22, demoSvolte: 14, offerte: 12, chiusure: 5, valoreContratto: 58500, incassato: 33500, spesa: 4200, slot: 50, oreContatto: 9 },
    { id: "v3", clienteId: "z2", periodo: mese(oggi()), lead: 74, chiamate: 168, discoveryFissate: 31, discoverySvolte: 17, demoFissate: 14, demoSvolte: 8, offerte: 7, chiusure: 0, valoreContratto: 0, incassato: 0, spesa: 3604.8, slot: 50, oreContatto: 21 },
    { id: "v4", clienteId: "z1", periodo: mese(piu(-31)), lead: 180, chiamate: 96, discoveryFissate: 42, discoverySvolte: 33, demoFissate: 24, demoSvolte: 19, offerte: 17, chiusure: 9, valoreContratto: 27400, incassato: 24100, spesa: 7100, slot: 60, oreContatto: 3 },
    { id: "v5", clienteId: "z1", periodo: mese(oggi()), lead: 141, chiamate: 88, discoveryFissate: 36, discoverySvolte: 24, demoFissate: 18, demoSvolte: 12, offerte: 10, chiusure: 0, valoreContratto: 0, incassato: 0, spesa: 7596.3, slot: 40, oreContatto: 4 },
  ],
  chiusure: [
    { id: "w1", clienteId: "z2", data: piu(-9), nome: "Andrea Sacchi", azienda: "Officine Sacchi", email: "a.sacchi@offsacchi.it", telefono: "0721 334455", funnel: "form", campagnaId: "k3", adset: "Titolari 40-60 · Marche", annuncio: "Video impianto — versione lunga", closer: "Io", valoreContratto: 31000, incassato: 12000, note: "Saldo a consegna impianto." },
    { id: "w2", clienteId: "z2", data: piu(-3), nome: "Marta Coppola", azienda: "Adriatica Logistica", email: "m.coppola@adrlog.it", telefono: "071 998412", funnel: "vsl", campagnaId: "k4", adset: "Retargeting 30 giorni", annuncio: "Carosello casi studio", closer: "Io", valoreContratto: 7400, incassato: 7400, note: "Pagato in un'unica soluzione." },
    { id: "w3", clienteId: "z1", data: piu(-6), nome: "Enoteca Bertelli", azienda: "Bertelli srl", email: "info@bertelli.it", telefono: "071 445566", funnel: "webinar", campagnaId: "k1", adset: "Interessi vino · lookalike 2%", annuncio: "Reel degustazione", closer: "Io", valoreContratto: 9800, incassato: 4900, note: "Metà ora, metà alla consegna." },
  ],
  file: [
    ...PROCEDURE_BASE(),
    { id: "fl3", clienteId: "z2", titolo: "Strategia 2026 — Tecnoplast", categoria: "strategia", tipo: "link", url: "https://drive.google.com/", tag: "annuale", aggiornato: piu(-45), testo: "" },
    { id: "fl4", clienteId: "z1", titolo: "Brief creatività autunno", categoria: "brief", tipo: "link", url: "https://drive.google.com/", tag: "creatività", aggiornato: piu(-12), testo: "" },
  ],
  funnel: [
    {
      id: "f1", clienteId: "z2", nome: "Impianti — VSL e prenotazione", obiettivo: "appuntamento", stato: "attiva",
      campagnaId: "k3", spesa: 3604.8, valoreMedio: 11600, modello: "call",
      passi: [
        { id: "p1", nome: "Annuncio LinkedIn", tipo: "annuncio", visite: 240300, conversioni: 1452 },
        { id: "p2", nome: "Landing con VSL", tipo: "vsl", visite: 1452, conversioni: 386 },
        { id: "p3", nome: "Modulo di qualifica", tipo: "form", visite: 386, conversioni: 74 },
        { id: "p4", nome: "Calendario", tipo: "calendario", visite: 74, conversioni: 31 },
        { id: "p5", nome: "Discovery svolta", tipo: "chiamata", visite: 31, conversioni: 17 },
      ],
      pagina: {
        titolo: "Quanto ti costa ogni ora di fermo impianto?",
        sottotitolo: "Analisi gratuita di 20 minuti: ti mostriamo dove si perde produttività e quanto si può recuperare entro il trimestre.",
        punti: "Sopralluogo tecnico senza impegno\nStima del recupero in euro, non in percentuali\nRisposta entro 24 ore dalla richiesta",
        cta: "Prenota l'analisi gratuita",
        webhook: "",
      },
      note: "",
    },
  ],
  rilevazioni: [
    { id: "r1", data: piu(-14), clienteId: "z1", campagnaId: "k1", spesa: 2980.5, impression: 402100, click: 3810, conversioni: 68, ricavi: 8420, frequenza: 2.9 },
    { id: "r2", data: piu(-7), clienteId: "z1", campagnaId: "k1", spesa: 4610.2, impression: 601300, click: 5260, conversioni: 96, ricavi: 11740, frequenza: 3.6 },
    { id: "r3", data: piu(-14), clienteId: "z2", campagnaId: "k3", spesa: 1720.4, impression: 118600, click: 742, conversioni: 12, ricavi: 0, frequenza: 1.8 },
    { id: "r4", data: piu(-7), clienteId: "z2", campagnaId: "k3", spesa: 2640.9, impression: 178200, click: 1090, conversioni: 17, ricavi: 0, frequenza: 2.1 },
  ],
  annotazioni: [
    { id: "an1", clienteId: "z1", data: piu(-9), tipo: "creativita", campagnaId: "k1", testo: "Sostituite due creatività stanche con reel degustazione. La frequenza era a 3,8." },
    { id: "an2", clienteId: "z1", data: piu(-5), tipo: "budget", campagnaId: "k1", testo: "Alzato il giornaliero da 220 a 300 € per il picco di stagione." },
    { id: "an3", clienteId: "z2", data: piu(-11), tipo: "pubblico", campagnaId: "k3", testo: "Ristretto il pubblico ai soli titolari, tolte le mansioni tecniche." },
  ],
  esperimenti: [
    { id: "e1", clienteId: "z2", ipotesi: "L'angolo sul fermo macchina batte l'angolo sul prezzo", variabile: "Gancio dell'annuncio", metrica: "CPA", inizio: piu(-40), fine: piu(-12), stato: "vinto", risultato: "CPA da 138 € a 96 €", conclusione: "Sul manifatturiero il costo del fermo funziona meglio dello sconto. Da usare come base." },
    { id: "e2", clienteId: "z1", ipotesi: "La landing con le recensioni in alto converte meglio", variabile: "Landing page", metrica: "Tasso di conversione", inizio: piu(-18), fine: "", stato: "corso", risultato: "", conclusione: "" },
  ],
  creativi: [
    { id: "d1", clienteId: "z1", campagnaId: "k1", tipo: "adset", nome: "Interessi vino · lookalike 2%", stato: "attiva", stadio: "1", spesa: 3980.2, impression: 512300, click: 4210, conversioni: 71, ricavi: 8420.5, frequenza: 4.61, note: "" },
    { id: "d2", clienteId: "z1", campagnaId: "k1", tipo: "annuncio", nome: "Reel degustazione", stato: "attiva", stadio: "1", spesa: 2140.6, impression: 268400, click: 2380, conversioni: 39, ricavi: 5100, frequenza: 3.98, note: "In calo dalla scorsa settimana." },
    { id: "d3", clienteId: "z1", campagnaId: "k1", tipo: "annuncio", nome: "Carosello cantina", stato: "pausa", stadio: "4", spesa: 291.6, impression: 31650, click: 313, conversioni: 8, ricavi: 718, frequenza: 2.14, note: "" },
    { id: "d4", clienteId: "z2", campagnaId: "k3", tipo: "adset", nome: "Titolari 40-60 · Marche", stato: "attiva", stadio: "1", spesa: 2410.4, impression: 158200, click: 980, conversioni: 15, ricavi: 0, frequenza: 2.62, note: "" },
    { id: "d5", clienteId: "z2", campagnaId: "k3", tipo: "annuncio", nome: "Video impianto — versione lunga", stato: "attiva", stadio: "3", spesa: 1194.4, impression: 82100, click: 472, conversioni: 6, ricavi: 0, frequenza: 2.11, note: "" },
    { id: "d6", clienteId: "z2", campagnaId: "k4", tipo: "adset", nome: "Retargeting 30 giorni", stato: "attiva", stadio: "4", spesa: 981.3, impression: 96400, click: 1902, clickLink: 1240, lead: 14, conversioni: 14, ricavi: 0, frequenza: 3.12, note: "" },
  ],
  automazioni: REGOLE_BASE(),
  registro: [],
  eseguiti: [],
  backup: { auto: "si", ultimoAuto: "", ultimoFile: "", errore: "" },
  snapshots: [],
  profili: [
    { id: "pf1", nome: "Meta Ads — campagne", tipo: "campagne", clienteId: "", url: "", mappa: {}, ultimo: "" },
  ],
});

const migra = (v) => ({
  agenzia: { nome: "NC", sottotitolo: "Media buying", logo: "" },
  aree: AREE_DEFAULT,
  clienti: (v.clienti || []).map((z) => ({ ...z, logo: "" })),
  contatti: v.contatti || [],
  trattative: v.trattative || [],
  campagne: v.campagne || [],
  task: (v.attivita || []).map((a) => ({
    id: a.id, clienteId: a.clienteId || "", areaId: "ar1", titolo: a.titolo, note: "",
    stato: a.fatto ? "fatto" : "dafare", priorita: "media", inizio: a.scadenza, scadenza: a.scadenza,
    stima: 0, ricorrenza: "no", checklist: [], creato: oggi(),
  })),
  vendite: [],
  chiusure: [],
  creativi: [],
  rilevazioni: [],
  periodi: [],
  giornate: [],
  deposito: [],
  modelli: [],
  eventi: [],
  categorie: CATEGORIE_BASE(),
  annotazioni: [],
  esperimenti: [],
  funnel: [],
  file: [],
  automazioni: [],
  registro: [],
  eseguiti: [],
  backup: { auto: "si", ultimoAuto: "", ultimoFile: "", errore: "" },
  snapshots: [],
});

/* I dati salvati con versioni precedenti possono non avere tutte le raccolte */
const normalizza = (d) => ({
  agenzia: { costoOrario: 0, lingua: "it", valuta: "EUR", ...(d.agenzia || { nome: "NC", sottotitolo: "Media buying", logo: "" }) },
  aree: d.aree && d.aree.length ? d.aree : AREE_DEFAULT,
  clienti: (d.clienti || []).map((z) => ({ margineLordo: 0, resi: 0, aov: 0, aoa: {}, padreId: "", archiviato: "no", valuta: "", valutaFee: "", ...z })),
  contatti: d.contatti || [], trattative: d.trattative || [],
  /* Il budget giornaliero non si deduce da un totale diviso la durata: o è quello
     impostato in piattaforma, o non c'è. Un numero inventato qui produce confronti falsi. */
  campagne: (d.campagne || []).map((k) => ({ archiviata: "no", erroreDal: "", modificaVista: "no", esitoManuale: "", storicoBudget: [], budgetGiorno: 0, ...k, bmk: { ...BMK_VUOTO, ...(k.bmk || {}) } })), task: d.task || [], chiusure: d.chiusure || [], creativi: (d.creativi || []).map((c) => ({ adsetId: "", ...c })),
  rilevazioni: d.rilevazioni || [],
  giornate: d.giornate || [],
  deposito: d.deposito || [],
  modelli: d.modelli || [],
  eventi: d.eventi || [],
  categorie: (d.categorie && d.categorie.length) ? d.categorie : CATEGORIE_BASE(),
  periodi: (d.periodi || []).map((pz) => ({ ...pz, tipo: { test1: "test", test2: "matura", settimana: "attiva" }[pz.tipo] || pz.tipo })),
  annotazioni: d.annotazioni || [], esperimenti: d.esperimenti || [], funnel: d.funnel || [],
  file: d.file === undefined ? PROCEDURE_BASE() : d.file,
  automazioni: d.automazioni || [], registro: d.registro || [], eseguiti: d.eseguiti || [],
  backup: { auto: "si", ultimoAuto: "", ultimoFile: "", errore: "", ...(d.backup || {}) },
  snapshots: d.snapshots || [],
  profili: d.profili || [],
  aiKey: d.aiKey || "",
  integrazioni: d.integrazioni || [],
  academy: { nome: "Accademia dell'Advertising", url: "https://accademiadelladvertising.applicazione.app/", urlAcquisto: "", incorpora: false, lezioni: [], corsi: [], viste: {}, ...(d.academy || {}) },
  vendite: (d.vendite || []).map((r) => {
    if (r.valoreContratto !== undefined) return r;
    const { valore = 0, ...resto } = r;           /* prima esisteva un solo importo */
    return { ...resto, valoreContratto: valore, incassato: 0 };
  }),
});

/* ================================================================== */
/* Numeri                                                              */
/* ================================================================== */

/* Legge un numero scritto in qualunque convenzione: 1.234,56 e 1,234.56 sono
   lo stesso valore. Decide l'ultimo separatore incontrato, che è sempre il decimale. */
const parseNum = (s) => {
  if (typeof s === "number") return isFinite(s) ? s : 0;
  if (s === null || s === undefined) return 0;
  let t = String(s).trim().replace(/[\s\u00A0'’]/g, "").replace(/[€$£¥%]/g, "");
  const ultimaV = t.lastIndexOf(","), ultimoP = t.lastIndexOf(".");
  if (ultimaV > -1 && ultimoP > -1) {
    /* Due separatori: quello più a destra è il decimale, l'altro è di migliaia */
    t = ultimaV > ultimoP
      ? t.replace(/\./g, "").replace(",", ".")
      : t.replace(/,/g, "");
  } else if (ultimaV > -1) {
    /* Solo virgole. Un'unica virgola con tre cifre dopo è ambigua ("2,345"):
       chi scrive in italiano intende un decimale, chi scrive in inglese le migliaia. */
    const dopo = t.length - ultimaV - 1;
    const inglese = LINGUA === "en";
    t = (t.split(",").length === 2 && (dopo !== 3 || !inglese)) ? t.replace(",", ".") : t.replace(/,/g, "");
  } else if (ultimoP > -1) {
    /* Solo punti: in italiano "1.500" sono millecinquecento, in inglese uno virgola cinque */
    const dopo = t.length - ultimoP - 1;
    if (t.split(".").length > 2 || (dopo === 3 && LINGUA !== "en")) t = t.replace(/\./g, "");
  }
  const n = parseFloat(t);
  return isFinite(n) ? n : 0;
};
const perInput = (n) => (n === 0 || n ? String(n).replace(".", ",") : "");
/* useGrouping: true mette sempre il separatore delle migliaia. Senza, l'italiano
   (come lo spagnolo) lo omette sotto le cinque cifre: 7217 invece di 7.217. */
const fmt = (n, d = 2) => new Intl.NumberFormat(LOCALE, { minimumFractionDigits: d, maximumFractionDigits: d, useGrouping: true }).format(n || 0);
const int = (n) => new Intl.NumberFormat(LOCALE, { useGrouping: true }).format(Math.round(n || 0));
const eur = (n, d = 2) => new Intl.NumberFormat(LOCALE, { style: "currency", currency: VALUTA, minimumFractionDigits: d, maximumFractionDigits: d, useGrouping: true }).format(n || 0);
/* Importi sempre con due decimali, anche nei totali: scelta del proprietario per uniformità */
const eur0 = (n) => eur(n, 2);
const pctv = (n, d = 2) => `${fmt(n, d)}%`;
const dataItA = (s) => (s ? new Date(s + "T00:00:00").toLocaleDateString(LOCALE, { day: "2-digit", month: "short", year: "numeric" }) : "—");
const dataIt = (s) => (s ? new Date(s + "T00:00:00").toLocaleDateString(LOCALE, { day: "2-digit", month: "short" }) : "—");
/* ==================================================================
/* La lingua corrente, tenuta fuori dal componente perché serve anche alle
   funzioni di modulo come et(). Il componente la riallinea a ogni disegno. */
let LINGUA = "it";
const impostaLingua = (l) => {
  const x = LINGUE.find((y) => y.v === l) || LINGUE[0];
  LINGUA = x.v; LOCALE = x.loc;
};

/* Traduce una stringa dell'interfaccia cercandola fra i valori italiani del dizionario:
   così gli elenchi di etichette restano scritti una volta sola. */
/* Numeri, date e valuta seguono la lingua scelta. Scrivere 1.234,56 a un inglese
   significa fargli leggere "uno virgola due": non è un dettaglio estetico. */
let LOCALE = "it-IT";
let VALUTA = "EUR";
const impostaValuta = (c) => { VALUTA = VALUTE.includes(c) ? c : "EUR"; };
/* La valuta di un cliente: la sua, altrimenti quella del cliente madre (risalendo),
   altrimenti quella dell'agenzia. Un sottocliente senza valuta segue la madre. */
const valutaCliente = (data, id) => {
  const clienti = (data && data.clienti) || [];
  const visti = new Set();
  let z = clienti.find((x) => x.id === id);
  while (z && !visti.has(z.id)) {
    if (z.valuta) return z.valuta;
    visti.add(z.id);
    z = z.padreId ? clienti.find((x) => x.id === z.padreId) : null;
  }
  return (data && data.agenzia && data.agenzia.valuta) || "EUR";
};
/* La valuta del compenso: un cliente può pagarti in una valuta diversa da quella
   delle sue campagne (campagne in USD, compenso in EUR). Se non è indicata, è la sua. */
const valutaCompenso = (data, id) => {
  const z = ((data && data.clienti) || []).find((x) => x.id === id);
  return (z && z.valutaFee) || valutaCliente(data, id);
};
/* Esegue fn con la valuta indicata e poi rimette quella di prima. Serve dove si
   scrivono importi di un cliente preciso fuori dalla sua scheda: avvisi, semaforo,
   elenchi di tutti i clienti. Vale per ciò che fn calcola subito, non per i
   componenti che React disegna dopo: a quelli la valuta va passata. */
const conValuta = (c, fn) => {
  const prima = VALUTA;
  impostaValuta(c);
  try { return fn(); } finally { VALUTA = prima; }
};
/* Per i componenti che ricevono il formattatore (grafici): fissa la valuta di adesso */
const fissaValuta = (f) => { const v = VALUTA; return (n) => conValuta(v, () => f(n)); };

/* en-gb non ha un dizionario proprio: usa quello inglese */
const base = (l) => (DIZ[l] ? l : String(l).split("-")[0]);

/* Formatta un importo nella valuta di quel cliente, non in quella attiva:
   serve negli elenchi dove convivono clienti con valute diverse. */
const eurDi = (n, valuta, d = 2) => new Intl.NumberFormat(LOCALE, {
  style: "currency", currency: VALUTE.includes(valuta) ? valuta : VALUTA,
  minimumFractionDigits: d, maximumFractionDigits: d, useGrouping: true,
}).format(n || 0);
const eurDi0 = (n, valuta) => eurDi(n, valuta, 2);

const T = (s) => {
  if (LINGUA === "it" || !s) return s;
  const tab = DIZ[base(LINGUA)] || {};
  const chiave = Object.keys(DIZ.it).find((k) => DIZ.it[k] === s);
  return chiave && tab[chiave] ? tab[chiave] : s;
};

const et = (arr, v) => T((arr.find((x) => x.v === v) || {}).l || "—");

/* ================================================================== */
/* Campagne: calcoli e avvisi                                          */
/* ================================================================== */

const calc = (k) => {
  const imp = k.impression || 0, cl = k.click || 0, sp = k.spesa || 0, cv = k.conversioni || 0, rc = k.ricavi || 0;
  const clk = k.clickLink || 0;                 /* clic sul link, non tutti i clic */
  const ld = k.lead || (k.obiettivo === "lead" ? cv : 0);
  return {
    imp, cl, clk, sp, cv, rc, ld, freq: k.frequenza || 0,
    ctr: imp ? (cl / imp) * 100 : 0, cpc: cl ? sp / cl : 0, cpm: imp ? (sp / imp) * 1000 : 0,
    /* Le metriche del link sono quelle su cui si giudica: se mancano, ripiego sui clic totali */
    ctrLink: imp ? ((clk || cl) / imp) * 100 : 0,
    cpcLink: (clk || cl) ? sp / (clk || cl) : 0,
    cpl: ld ? sp / ld : 0,
    cpa: cv ? sp / cv : 0, roas: sp ? rc / sp : 0, cr: cl ? (cv / cl) * 100 : 0, usato: k.budget ? sp / k.budget : 0,
    soloTotali: clk === 0 && cl > 0,
  };
};

const quotaPeriodo = (k) => {
  if (!k.inizio || !k.fine) return null;
  const i = new Date(k.inizio), f = new Date(k.fine), o = new Date(oggi());
  const tot = (f - i) / 864e5;
  if (tot <= 0) return null;
  return Math.min(1, Math.max(0, (o - i) / 864e5 / tot));
};

/* Le note non sono problemi: non entrano nei conteggi degli avvisi */
/* Ordinamento per cognome: tolgo titoli e forme societarie, tengo le particelle
   attaccate al cognome (De Luca, Di Marco) e uso l'ultima parola rimasta. */
const TITOLI = ["avv", "avvocato", "dott", "dottor", "dottore", "dottssa", "dottoressa", "dr", "prof",
  "ing", "arch", "geom", "rag", "sig", "sigra", "studio", "legale", "associato", "associati", "st"];
const CODE = ["srl", "srls", "spa", "sas", "snc", "sr", "stp", "ssd", "partners", "partner", "associati",
  "company", "group", "co", "ltd", "inc"];
const PARTICELLE = ["de", "di", "del", "della", "dello", "dei", "degli", "da", "dal", "dalla", "la", "lo",
  "li", "van", "von", "der", "den", "mc", "mac", "o", "san", "santa"];

const chiaveCognome = (nome, cognome) => {
  if ((cognome || "").trim()) return cognome.trim().toLowerCase();
  const parole = String(nome || "")
    .replace(/&/g, " ")
    .split(/[\s.,]+/)
    .map((x) => x.trim())
    .filter(Boolean)
    .filter((x) => !TITOLI.includes(x.toLowerCase().replace(/[^a-zà-ù]/g, "")));
  let p = parole.filter((x) => !CODE.includes(x.toLowerCase().replace(/[^a-z]/g, "")));
  if (!p.length) p = parole.length ? parole : [String(nome || "")];
  const ultima = p[p.length - 1];
  const prec = p.length > 1 ? p[p.length - 2] : "";
  const composto = prec && PARTICELLE.includes(prec.toLowerCase()) ? `${prec} ${ultima}` : ultima;
  return composto.toLowerCase();
};

const perCognome = (a, b) =>
  chiaveCognome(a.nome, a.cognome).localeCompare(chiaveCognome(b.nome, b.cognome), "it") ||
  String(a.nome || "").localeCompare(String(b.nome || ""), "it");


const veri = (l) => (l || []).filter((a) => a.gravita !== "nota");

/* Motore di valutazione. Segue l'ordine del framework: prima si controlla se
   c'è abbastanza spesa per giudicare, poi si legge il resto. */
function avvisi(k, ctx = {}) {
  const d = calc(k), b = k.bmk || {}, out = [];
  /* Errore di pagamento: la campagna è ferma e ogni giorno conta.
     L'avviso resta finché non cambi stato, e non si mescola con le valutazioni di rendimento. */
  if (k.stato === "errore" && k.archiviata !== "si") {
    const g = k.erroreDal ? giorniTra(k.erroreDal, oggi()) : 0;
    const cf = causaFermo(k);
    return [{
      campagnaId: k.id, clienteId: k.clienteId, gravita: "critica",
      metrica: k.causaErrore && k.causaErrore !== "pagamento" ? "Account" : "Pagamento", valore: g ? `${int(g)} ${g === 1 ? "giorno" : "giorni"}` : "in corso",
      soglia: "da risolvere",
      titolo: cf.titolo,
      cosa: (k.erroreDal
        ? `Segnalata il ${dataItA(k.erroreDal)}: sono ${int(g)} ${g === 1 ? "giorno" : "giorni"} che non eroga.`
        : `La campagna è segnata come "${cf.l}" e non sta erogando.`) + (k.accountBloccato ? ` Stato dell'account su Meta: ${k.accountBloccato}.` : ""),
      causa: cf.causa,
      azione: cf.azione,
    }];
  }

  /* Una campagna spenta senza data di fine lascia i conti aperti: lo dico prima
     di uscire, perché per gli stati chiusi il resto della valutazione non gira. */
  if (["conclusa", "pausa"].includes(k.stato) && !k.fine && k.archiviata !== "si" && (k.spesa || 0) > 0) {
    return [{
      campagnaId: k.id, clienteId: k.clienteId, gravita: "nota",
      metrica: "Date", valore: "fine assente", soglia: "da compilare",
      titolo: "Manca la data di fine",
      cosa: `La campagna è in stato "${et(STATI_CAMPAGNA, k.stato)}" ma non ha una data di chiusura.`,
      causa: "Senza quella data l'app continua a contare i giorni fino a oggi, e il periodo della campagna cresce su giorni in cui non hai speso.",
      azione: "Aprila e compila la data di fine: i conteggi si fermano lì e i confronti tornano corretti.",
    }];
  }
  if (["conclusa", "pausa"].includes(k.stato) || k.archiviata === "si") return out;
  /* Senza spesa registrata non c'è niente da valutare: sarebbero allarmi su dati mancanti */
  if (d.sp <= 0) return out;
  const F = faseCampagna(k);
  /* In fase Matura nessun giudizio è definitivo: resta un'attenzione,
     tranne dove il problema è tecnico o economico e non di rendimento. */
  const tecnici = ["Conversioni", "Erogazione", "Margine"];
  /* Cliente in attivo: se il Ciclo di vendita dice che i contratti superano la spesa (oltre il
     pareggio, quando il margine è noto), i numeri della singola campagna non sono un allarme.
     Restano segnalati, ma come attenzione e con il motivo; critici solo i fermi (pagamento, account). */
  const R = ctx.redditivita;
  const push = (o) => {
    const matura = F.fase === "matura" && o.gravita === "critica" && !tecnici.includes(o.metrica);
    const inAttivo = R && o.gravita === "critica";
    out.push({
      ...o, campagnaId: k.id, clienteId: k.clienteId,
      gravita: matura || inAttivo ? "attenzione" : o.gravita,
      ...(inAttivo ? { cosa: `${o.cosa} Il cliente però è in attivo: negli ultimi ${R.mesi} ${R.mesi === 1 ? "mese" : "mesi"} del Ciclo di vendita i contratti valgono ${fmt(R.ritorno)} volte la spesa, quindi non è un allarme ma un margine di miglioramento.`, inAttivo: true } : {}),
    });
  };
  const z = ctx.cliente;
  const iv = intervallo(k, ctx.prec);
  /* Su una campagna matura si giudica la finestra recente, non la media da sempre */
  const W = F.fase === "attiva" ? ctx.finestra : null;
  const V = W || d;
  const dove = W ? `negli ultimi ${W.giorni} giorni` : "dall'inizio della campagna";

  /* --- erogazione: spesa contro atteso, sullo stesso intervallo di date --- */
  const bg = budgetDelGiorno(k, oggi());
  const cambi = (k.storicoBudget || []).length;
  const E = ctx.erogazione;
  if (E && (E.sospetto || E.storte > 0)) push({
    gravita: "attenzione",
    metrica: "Date", valore: E.storte ? `${E.storte} finestre` : `${int(E.giorni)} giorni`, soglia: "da correggere",
    titolo: "Le date di questa campagna non tornano",
    cosa: E.storte
      ? `${E.storte} finestre hanno date impossibili — più lunghe di due mesi o con la fine prima dell'inizio — e le ho escluse dal calcolo.`
      : `La campagna risulta partita il ${dataItA(E.dal)} e ancora attiva: ${int(E.giorni)} giorni. Con una finestra così lunga il confronto tra speso e previsto non significa niente, quindi l'ho sospeso.`,
    causa: "Quasi sempre è l'anno sbagliato nella data di partenza o in una finestra generata prima di correggerla.",
    azione: "Apri la scheda, controlla la data di partenza e le date delle fasi. Se sono sbagliate, cancella le finestre e usa \"Genera le finestre\": le ricostruisce con i tagli giusti.",
  });
  else if (E && E.senzaFinestre) push({
    gravita: "nota",
    metrica: "Erogazione", valore: `${int(E.giorni)} giorni`, soglia: "servono le finestre",
    titolo: "Non posso valutare il ritmo di spesa",
    cosa: `La campagna gira da ${int(E.giorni)} giorni e i dati che hai inserito sono un totale unico: non so a quali giorni si riferiscono.`,
    causa: "Su una campagna lunga il totale complessivo non dice nulla sul ritmo attuale.",
    azione: "Apri il pulsante Fasi e genera le finestre: da lì il confronto tra speso e previsto torna a essere sensato.",
  });
  /* Erogazione: si confronta quanto spende davvero in un giorno con quanto
     dovrebbe spenderne. Nessun totale, nessuna proiezione: due numeri omogenei. */
  /* Con date incoerenti o senza finestre il confronto non ha senso: l'avviso sulle date basta */
  if (bg > 0 && E && E.giorni >= 3 && k.stato === "attiva" && !E.sospetto && !E.senzaFinestre && !(E.storte > 0)) {
    const spesaGiorno = d.sp / E.giorni;
    const attesoGiorno = E.atteso > 0 ? E.atteso / E.giorni : bg;
    const quota = attesoGiorno > 0 ? spesaGiorno / attesoGiorno : 0;
    const suQuali = E.daFasi ? `nelle finestre compilate (${dataItA(E.dal)} – ${dataItA(E.al)})` : `dal ${dataItA(E.dal)} a oggi`;
    const confronto = `${eur(spesaGiorno)} al giorno contro ${eur(attesoGiorno)} impostati, su ${E.giorni} giorni ${suQuali}`;

    /* Prima di gridare al sotto-consumo: il budget scritto in scheda è credibile? */
    if (quota < 0.1) push({
      gravita: "attenzione",
      metrica: "Budget", valore: `${eur(spesaGiorno)}/giorno`, soglia: `${eur(attesoGiorno)}/giorno`,
      titolo: "Il budget in scheda non corrisponde alla spesa reale",
      cosa: `Questa campagna spende ${eur(spesaGiorno)} al giorno, ma in scheda risulta un budget di ${eur(attesoGiorno)}: ${fmt(attesoGiorno / Math.max(spesaGiorno, 0.01), 0)} volte tanto.`,
      causa: "Di solito è un budget totale finito nel campo del giornaliero, o una cifra inserita con un separatore diverso.",
      azione: "Apri la campagna e correggi il budget giornaliero con quello vero della piattaforma. Finché i due numeri non sono omogenei, ogni confronto sull'erogazione è privo di senso.",
    });
    else if (quota > 1.2) push({
      gravita: quota > 1.5 ? "critica" : "attenzione",
      metrica: "Erogazione", valore: `${eur(spesaGiorno)}/giorno`, soglia: `${eur(attesoGiorno)}/giorno`,
      titolo: "Spende più del budget giornaliero impostato",
      cosa: `${confronto}${cambi > 1 ? `, tenendo conto dei ${cambi} cambi di budget` : ""}: ${fmt(quota * 100, 0)}% del previsto.`,
      causa: "Budget alzato in piattaforma e non registrato qui, oppure più ad set che spingono sullo stesso budget di campagna.",
      azione: "Allinea il numero in scheda a quello vero: se il giornaliero è stato alzato apposta, aggiungi la variazione nello storico.",
    });
    else if (quota < 0.75) push({
      gravita: "attenzione",
      metrica: "Erogazione", valore: `${eur(spesaGiorno)}/giorno`, soglia: `${eur(attesoGiorno)}/giorno`,
      titolo: "Il budget giornaliero non viene consumato",
      cosa: `${confronto}: eroga il ${fmt(quota * 100, 0)}% di quello che potrebbe.`,
      causa: "Offerta troppo bassa per l'asta, pubblico troppo stretto, limitazioni di posizionamento — oppure le finestre compilate non coprono tutti i giorni in cui ha girato.",
      azione: "Controlla prima che le finestre corrispondano ai giorni reali. Se i dati sono giusti, alza il limite di offerta o allarga pubblico e posizionamenti.",
    });
  }

  /* --- i tre giorni sono scaduti: ora la decisione va presa --- */
  if (k.ultimaModifica && k.modificaVista !== "si" && giorniTra(k.ultimaModifica, oggi()) > ATTESA_MODIFICA) {
    const scaduta = sommaGiorni(k.ultimaModifica, ATTESA_MODIFICA);
    const da = giorniTra(scaduta, oggi()) - 1;
    push({
      gravita: da > 2 ? "attenzione" : "nota",
      metrica: "Valutazione", valore: da > 0 ? `da ${int(da)} ${da === 1 ? "giorno" : "giorni"}` : "da oggi", soglia: dataIt(scaduta),
      titolo: "Valutazione pronta dopo la modifica",
      cosa: `I tre giorni di attesa dalla modifica del ${dataItA(k.ultimaModifica)}${k.notaModifica ? ` — ${k.notaModifica}` : ""} sono scaduti il ${dataItA(scaduta)}.`,
      causa: "È il momento in cui l'intervento ha prodotto abbastanza dati puliti per dire se ha funzionato.",
      azione: "Apri il Semaforo, guarda il segnale e decidi. Poi premi “Ho rivisto” per spegnere la campanella: finché resta accesa, quella decisione è ancora da prendere.",
    });
  }

  /* --- dopo un intervento, tre giorni di sospensione del giudizio --- */
  if (k.ultimaModifica && giorniTra(k.ultimaModifica, oggi()) <= ATTESA_MODIFICA) {
    const quando = sommaGiorni(k.ultimaModifica, ATTESA_MODIFICA);
    push({
      gravita: "nota",
      metrica: "Attesa", valore: `${int(giorniTra(k.ultimaModifica, oggi()) - 1)} di ${ATTESA_MODIFICA} giorni`, soglia: dataIt(quando),
      titolo: "Modifica recente: giudizio sospeso",
      cosa: `È stata modificata il ${dataItA(k.ultimaModifica)}${k.notaModifica ? ` — ${k.notaModifica}` : ""}.`,
      causa: "Dopo un intervento l'algoritmo ritara la consegna: i numeri dei primi giorni descrivono la transizione, non la campagna.",
      azione: `Rivedila il ${dataItA(quando)}. Fino ad allora non toccare altro, o non saprai cosa ha prodotto l'effetto.`,
    });
  }

  /* --- fase 1: i primi giorni non si giudicano --- */
  if (F.fase === "test") {
    push({
      gravita: "nota",
      metrica: "Finestra", valore: `giorno ${F.g}`, soglia: `dal giorno ${TEST_1 + 1}`,
      titolo: "Primo test in corso: non decidere adesso",
      cosa: `La campagna è partita il ${dataIt(k.inizio)} e ha ${F.g} ${F.g === 1 ? "giorno" : "giorni"} di vita. Ha speso ${eur(d.sp)}.`,
      causa: "Nei primi tre giorni la consegna è ancora instabile: quello che vedi è la fase di apprendimento, non il rendimento della campagna.",
      azione: `Annota i numeri e non toccare niente. Dal quarto giorno entra nella fase Matura, e dal giorno ${TEST_2 + 1} diventa Attiva e si giudica sulla finestra di ${FINESTRA} giorni.`,
    });
    return out;
  }

  /* --- 1. soglia di spesa: sotto questa cifra ogni giudizio è rumore --- */
  const aov = z && z.aov ? z.aov : 0;
  /* Il costo obiettivo: nelle campagne a contatti è il CPL, altrove il CPA */
  const tetto = kpiCosto(k, d);
  const soglia = k.obiettivo === "vendite" ? aov * 2 : (tetto > 0 ? tetto * 2 : 0);
  if (soglia > 0 && d.sp < soglia) {
    push({
      gravita: "nota",
      metrica: "Soglia", valore: eur(d.sp), soglia: `≥ ${eur(soglia)}`,
      titolo: "Troppo presto per giudicare",
      cosa: `Questa campagna ha speso ${eur(d.sp)} contro una soglia di valutazione di ${eur(soglia)}, cioè due volte ${k.obiettivo === "vendite" ? "il valore medio dell'ordine" : suiLead(k, d) ? "il CPL obiettivo" : "il CPA obiettivo"}.`,
      causa: "Sotto quella cifra l'algoritmo non ha ancora abbastanza segnale per ottimizzare la consegna: quello che vedi è rumore, non un risultato.",
      azione: "Annota la tendenza e non toccare niente. Torna a decidere quando la soglia è raggiunta: fermare adesso significa scartare un'ipotesi che non hai ancora testato.",
    });
    return out;
  }

  /* --- 2. CTR: se il gancio non ferma lo scorrimento è un problema di creatività --- */
  if (b.ctrMin > 0 && V.imp > 1000 && (V.ctrLink || V.ctr) < b.ctrMin) {
    push({
      gravita: (V.ctrLink || V.ctr) < b.ctrMin * 0.6 ? "critica" : "attenzione",
      metrica: "CTR", valore: pctv(V.ctrLink || V.ctr), soglia: `≥ ${pctv(b.ctrMin)}`,
      titolo: "Il gancio non funziona",
      cosa: `${pctv(V.ctrLink || V.ctr)} di clic sul link su ${int(V.imp)} impression ${dove}, con ${eur(V.sp)} di spesa nel periodo valutato.${V.soloTotali ? " Attenzione: qui stai usando i clic totali, non quelli sul link." : ""}`,
      causa: "CTR basso a spesa sufficiente significa che la creatività non ferma lo scorrimento. È un problema di creatività, non di pubblico: cambiare targeting qui è l'errore classico.",
      azione: "Spegni questa creatività e tieni lo stesso pubblico. Riscrivi il gancio con un angolo diverso, mantenendo la stessa offerta.",
    });
  }

  /* --- 3. CPM contro la media dell'account: salute dell'asta --- */
  const rifCpm = ctx.cpmMedio > 0 ? ctx.cpmMedio * 1.4 : (b.cpmMax > 0 ? b.cpmMax : 0);
  const cpmV = W ? (W.imp ? (W.sp / W.imp) * 1000 : 0) : d.cpm;
  if (rifCpm > 0 && cpmV > rifCpm) {
    push({
      gravita: ctx.cpmMedio > 0 && cpmV > ctx.cpmMedio * 1.8 ? "critica" : "attenzione",
      metrica: "CPM", valore: eur(cpmV), soglia: `≤ ${eur(rifCpm)}`,
      titolo: "Stai pagando l'asta più degli altri",
      cosa: ctx.cpmMedio > 0
        ? `${eur(cpmV)} ${dove}, contro una media dell'account di ${eur(ctx.cpmMedio)}: oltre il 40% in più.`
        : `${eur(cpmV)} contro il riferimento di ${eur(rifCpm)} che hai impostato.`,
      causa: "Un CPM alto in rapporto al tuo account significa che l'algoritmo paga un sovrapprezzo per consegnare l'annuncio, perché non riceve abbastanza segnali positivi per vincere l'asta in modo efficiente. La creatività non convince.",
      azione: "Sostituisci la creatività lasciando lo stesso ad set: il pubblico può essere ancora valido, prova un annuncio nuovo lì dentro prima di spostare il targeting.",
    });
  }

  /* --- 4. cancello ROAS / CPA, con la clausola sulla tendenza --- */
  const migliora = iv && ((d.rc > 0 && iv.rc > 0 && iv.roas > d.roas) || (iv.cv > 0 && d.cv > 0 && iv.cpa < d.cpa));
  if (b.roasMin > 0 && V.rc > 0 && V.roas < b.roasMin) {
    push({
      gravita: migliora ? "attenzione" : "critica",
      metrica: "ROAS", valore: `${fmt(V.roas)}×`, soglia: `≥ ${fmt(b.roasMin)}×`,
      titolo: migliora ? "Sotto obiettivo, ma in ripresa" : "Sotto obiettivo e senza ripresa",
      cosa: migliora
        ? `ROAS a ${fmt(V.roas)}× ${dove} contro un obiettivo di ${fmt(b.roasMin)}×, però nell'ultimo intervallo rilevato è a ${fmt(iv.roas)}×: sta risalendo.`
        : `ROAS a ${fmt(V.roas)}× ${dove}, contro un obiettivo di ${fmt(b.roasMin)}×, senza segnali di miglioramento recente.`,
      causa: migliora
        ? "Un singolo periodo negativo è rumore. Una risalita in corso può essere l'algoritmo che esce da una finestra di consegna disturbata."
        : "Spesa sufficiente più assenza di tendenza in miglioramento è un fatto, non un caso: l'angolo non regge contro questo pubblico.",
      azione: migliora
        ? "Lascia correre ancora un intervallo prima di decidere, ma fissa la data in cui deciderai."
        : "Spegnila e scrivi nel Diario quale angolo ha fallito: non ritestare lo stesso gancio sullo stesso pubblico.",
    });
  }
  /* Il CPL ha il suo cancello: nelle campagne a contatti il lead è il primo punto
     di conversione, e se costa troppo il problema esiste già lì, prima del CPA. */
  const tettoL = tettoCpl(b);
  const cplGiu = iv && iv.ld > 0 && V.ld > 0 && iv.cpl > 0 && iv.cpl < V.cpl;
  if (tettoL > 0 && V.ld > 0 && V.cpl > tettoL) {
    push({
      gravita: cplGiu ? "attenzione" : V.cpl > tettoL * 1.4 ? "critica" : "attenzione",
      metrica: "CPL", valore: eur(V.cpl), soglia: `≤ ${eur(tettoL)}`,
      titolo: cplGiu ? "Costo per lead alto, ma in discesa" : "Costo per lead oltre il tetto",
      cosa: cplGiu
        ? `${eur(V.cpl)} per lead ${dove} contro un tetto di ${eur(tettoL)}, ma nell'ultimo intervallo sei a ${eur(iv.cpl)}.`
        : `${eur(V.cpl)} per lead ${dove}, contro un tetto di ${eur(tettoL)}: ${int(V.ld)} lead per ${eur(V.sp)}.${b.cplMax > 0 ? "" : " Il tetto è quello del CPA, perché il CPL obiettivo non è impostato."}`,
      causa: `Il clic sul link costa ${eur((V.clk || V.cl) ? V.sp / (V.clk || V.cl) : 0)} e diventa lead nel ${pctv((V.clk || V.cl) ? (V.ld / (V.clk || V.cl)) * 100 : 0)} dei casi. Se il costo del clic è in linea, il punto di caduta è il modulo o la pagina; se è salito, è l'asta o il pubblico.`,
      azione: cplGiu
        ? "Aspetta il prossimo intervallo prima di spegnere, ma fissa la data in cui deciderai."
        : "Lavora prima sul passaggio dal clic al lead — modulo più corto, promessa più chiara — e solo dopo sposta pubblico o budget. Il CPL è il primo punto dove la campagna si rompe.",
    });
  }
  /* Dove conversioni e lead sono lo stesso numero il CPA ripeterebbe il CPL: taccio */
  if (b.cpaMax > 0 && V.cv > 0 && V.cpa > b.cpaMax && !(suiLead(k, d) && V.cv === V.ld && tettoL > 0)) {
    push({
      gravita: migliora ? "attenzione" : V.cpa > b.cpaMax * 1.4 ? "critica" : "attenzione",
      metrica: "CPA", valore: eur(V.cpa), soglia: `≤ ${eur(b.cpaMax)}`,
      titolo: migliora ? "Costo alto, ma in discesa" : "Costo per acquisizione oltre il tetto",
      cosa: migliora
        ? `${eur(V.cpa)} ${dove} contro un tetto di ${eur(b.cpaMax)}, ma nell'ultimo intervallo sei a ${eur(iv.cpa)}.`
        : `${eur(V.cpa)} per conversione ${dove}, contro un tetto di ${eur(b.cpaMax)}: ${int(V.cv)} conversioni per ${eur(V.sp)}.`,
      causa: `Il click costa ${eur(V.cl ? V.sp / V.cl : 0)} e converte al ${pctv(V.cl ? (V.cv / V.cl) * 100 : 0)}. Se il CPC è in linea il problema è dopo il click; se è salito, è l'asta o il pubblico.`,
      azione: migliora
        ? "Aspetta il prossimo intervallo prima di spegnere, ma non oltre."
        : "Spegnila e documenta l'angolo che non ha funzionato. Se il CPC era in linea, il lavoro da fare è sulla pagina, non sul targeting.",
    });
  }

  /* --- 5. slope test: il calo lento, quello che non fa rumore --- */
  const serie = (ctx.serie || []).slice(-4);
  if (serie.length >= 3) {
    const punti = [];
    for (let i = 1; i < serie.length; i++) {
      const p = intervallo(serie[i], serie[i - 1]);
      if (p) punti.push(p);
    }
    if (iv) punti.push(iv);
    if (punti.length >= 3) {
      const ultimi = punti.slice(-3);
      const calaRoas = d.rc > 0 && ultimi.every((p, i) => i === 0 || (p.roas > 0 && p.roas < ultimi[i - 1].roas));
      const saleCpa = d.cv > 0 && ultimi.every((p, i) => i === 0 || (p.cpa > 0 && p.cpa > ultimi[i - 1].cpa));
      if (calaRoas || saleCpa) push({
        gravita: "critica",
        metrica: "Tendenza", valore: calaRoas ? `${fmt(ultimi[2].roas)}×` : eur(ultimi[2].cpa),
        soglia: calaRoas ? `era ${fmt(ultimi[0].roas)}×` : `era ${eur(ultimi[0].cpa)}`,
        titolo: "Declino controllato: la creatività sta morendo",
        cosa: `Tre intervalli consecutivi di peggioramento senza cause esterne evidenti: ${calaRoas ? `ROAS da ${fmt(ultimi[0].roas)}× a ${fmt(ultimi[2].roas)}×` : `CPA da ${eur(ultimi[0].cpa)} a ${eur(ultimi[2].cpa)}`}.`,
        causa: "È affaticamento al rallentatore. L'annuncio funzionava e l'algoritmo continua a consegnarlo perché lo storico è forte, ma l'interesse reale sta crollando adesso.",
        azione: "Limita il budget e lancia subito una creatività sfidante. Aspettare il fondo costa spesa ogni settimana che passa: l'annuncio attuale ti compra tempo, non recupero.",
      });
    }
  }

  /* --- 6. l'eccezione: frequenza alta ma risultati a target --- */
  const aTarget = (d.rc > 0 && b.roasMin > 0 && d.roas >= b.roasMin)
    || (d.ld > 0 && tettoL > 0 && d.cpl <= tettoL)
    || (d.cv > 0 && b.cpaMax > 0 && d.cpa <= b.cpaMax);
  if (b.freqMax > 0 && d.freq > b.freqMax) {
    if (aTarget) push({
      gravita: "nota",
      metrica: "Frequenza", valore: fmt(d.freq), soglia: `soglia ${fmt(b.freqMax)}`,
      titolo: "Frequenza alta, ma non spegnere",
      cosa: `La stessa persona ha visto l'annuncio ${fmt(d.freq)} volte, e i risultati sono ancora a obiettivo.`,
      causa: "Frequenza alta con risultati forti significa che la creatività ha una tenuta insolita: il mercato non se n'è stancato.",
      azione: "Lasciala correre e continua a controllare. Quasi tutti spengono per abitudine quando la frequenza tocca 3 o 4: se i numeri dicono il contrario, hanno ragione i numeri.",
    });
    else push({
      gravita: d.freq > b.freqMax * 1.3 ? "critica" : "attenzione",
      metrica: "Frequenza", valore: fmt(d.freq), soglia: `≤ ${fmt(b.freqMax)}`,
      titolo: "Pubblico saturo e risultati sotto obiettivo",
      cosa: `Frequenza a ${fmt(d.freq)} e risultati che non sono a target: stai ricomprando le stesse persone senza ottenere di più.`,
      causa: "Il pubblico è troppo piccolo per il budget erogato. Qui la frequenza non è un dettaglio: è la spiegazione del costo che sale.",
      azione: "Allarga il pubblico o abbassa il giornaliero, escludi chi ha già convertito e ruota le creatività.",
    });
  }

  /* --- tracciamento e pagina --- */
  /* Nelle campagne a lead il risultato è il lead: le conversioni di acquisto non arrivano mai
     da Meta e contarle dava un allarme critico a tutte le campagne a contatti */
  const aLead = k.obiettivo === "lead";
  if ((aLead ? V.ld : V.cv) === 0 && V.sp > 0) push({
    gravita: "critica",
    metrica: "Conversioni", valore: "0", soglia: "> 0",
    titolo: aLead ? "Spesa oltre soglia senza lead" : "Spesa oltre soglia senza conversioni",
    cosa: `${eur(V.sp)} spesi ${dove} e ${aLead ? "nessun lead registrato" : "nessuna conversione registrata"}, con ${int(V.cl)} click arrivati a destinazione.`,
    causa: "Superata la soglia di valutazione, zero conversioni non è più pazienza: o il tracciamento non registra, o l'offerta non regge.",
    azione: "Fai una conversione di test e controlla che arrivi. Se il tracciamento è a posto, spegni: è il caso in cui aspettare costa e basta.",
  });
  const crV = V.cl ? (V.cv / V.cl) * 100 : 0;
  if (V.cl > 300 && V.cv > 0 && crV < 1) push({
    gravita: "attenzione",
    metrica: "Conversione", valore: pctv(crV), soglia: "≥ 1,00%",
    titolo: "Il traffico arriva ma non converte",
    cosa: `${int(V.cl)} click ${dove} hanno prodotto ${int(V.cv)} conversioni: ${pctv(crV)}.`,
    causa: "Il problema è a valle dell'annuncio: pagina lenta su mobile, promessa diversa da quella dell'annuncio, modulo troppo lungo, prezzo scoperto tardi.",
    azione: "Allinea il titolo della pagina al testo dell'annuncio, togli campi al modulo e misura il caricamento da mobile. Qui si guadagna più che ritoccando il targeting.",
  });

  /* --- margine: il ROAS va letto contro il pareggio del cliente --- */
  const rp = roasPareggio(z);
  if (rp > 0 && d.rc > 0 && d.roas < rp) {
    const p = profitto(d.rc, d.sp, z);
    push({
      gravita: "critica",
      metrica: "Margine", valore: `${fmt(d.roas)}×`, soglia: `≥ ${fmt(rp)}×`,
      titolo: "Sotto il punto di pareggio",
      cosa: `Con un margine netto del ${fmt(margineNetto(z) * 100, 1)}% il pareggio è a ${fmt(rp)}×, e sei a ${fmt(d.roas)}×: su questa campagna il cliente perde ${eur(Math.abs(p))}.`,
      causa: "Il ROAS da solo nasconde il margine: si vende, ma ogni vendita costa più di quanto rende.",
      azione: "Prima di aumentare il budget riporta il ROAS sopra il pareggio: escludi i prodotti sotto margine e alza il valore medio dell'ordine invece del volume.",
    });
  }

  return out;
}

/* Margine: un ROAS ha senso solo confrontato con il punto di pareggio del cliente */
const margineNetto = (z) => (z && z.margineLordo ? (z.margineLordo / 100) * (1 - (z.resi || 0) / 100) : 0);
const roasPareggio = (z) => { const mn = margineNetto(z); return mn > 0 ? 1 / mn : 0; };
const profitto = (ricavi, spesa, z) => ricavi * margineNetto(z) - spesa;

/* Lo stesso semaforo applicato a un ad set o a un annuncio: stesse soglie,
   stesso KPI della campagna. Il ritmo di spesa resta un fatto di campagna,
   quindi qui si giudica solo il rapporto fra costo e obiettivo. */
const semaforoFiglio = (c, kpi, pesoCampagna) => {
  const d = calc(c);
  const conv = d.ld > 0 ? d.ld : d.cv;
  const costo = conv > 0 ? d.sp / conv : 0;
  const rapporto = costo ? (costo / kpi) * 100 : 0;
  const quota = pesoCampagna ? d.sp / pesoCampagna : 0;
  const base = { costo, rapporto, conv, sp: d.sp, quota, ctr: d.ctrLink || d.ctr };

  if (!kpi) return { ...base, segnale: "grigio", perche: "KPI non impostato sulla campagna." };
  if (["pausa", "conclusa"].includes(c.stato)) return { ...base, segnale: "grigio", perche: "Non in corso." };
  /* Sotto una volta il KPI di spesa il dato non regge: lo stesso principio della campagna */
  if (d.sp < kpi) return { ...base, segnale: "grigio", perche: `Ha speso ${eur(d.sp)}, meno di una volta il KPI: troppo poco per giudicarlo.` };
  if (conv === 0) return { ...base, segnale: "rosso", perche: `${eur(d.sp)} spesi senza una conversione.` };
  if (rapporto > 150) return { ...base, segnale: "rosso", perche: `${eur(costo)} per risultato, ${fmt(rapporto, 0)}% del KPI.` };
  if (rapporto > 100) return { ...base, segnale: "giallo", perche: `${eur(costo)} per risultato, ${fmt(rapporto, 0)}% del KPI.` };
  return { ...base, segnale: "verde", perche: `${eur(costo)} per risultato, ${fmt(rapporto, 0)}% del KPI.` };
};

/* Finestre di valutazione: una campagna giovane non si giudica come una matura.
   Giorni 1-3: fase Test, non si decide. Giorni 4-10: fase Matura. Dal giorno 11:
   fase Attiva, e l'analisi si fa sugli ultimi sette giorni, non sul totale da inizio campagna. */
const TEST_1 = 3, TEST_2 = 10, FINESTRA = 7;
/* Dopo un intervento si aspettano tre giorni prima di rigiudicare:
   l'algoritmo ha bisogno di ritarare la consegna, e i primi dati sono rumore. */
const ATTESA_MODIFICA = 3;

/* Calendario: le categorie danno il colore alle date. Sono modificabili,
   perché il significato dei colori lo decidi tu, non il software. */
const CATEGORIE_BASE = () => [
  { id: uid(), nome: "Lancio campagna", colore: "#D8455F" },
  { id: uid(), nome: "Revisione dati", colore: "#4C8DF6" },
  { id: uid(), nome: "Call con il cliente", colore: "#3FA45C" },
  { id: uid(), nome: "Scadenza", colore: "#C99A3A" },
  { id: uid(), nome: "Obiettivo", colore: "#8B5CF6" },
];

const VISTE_CAL = [
  { v: "giorno", l: "Giorno" }, { v: "settimana", l: "Settimana" },
  { v: "mese", l: "Mese" }, { v: "anno", l: "Anno" },
];

/* Cosa può comparire sul calendario, e con che colore quando non ha categoria */
const FONTI_CAL = [
  { v: "eventi", l: "Eventi e obiettivi", c: "#8B5CF6" },
  { v: "task", l: "Task", c: "#C99A3A" },
  { v: "campagne", l: "Avvii e chiusure campagne", c: "#D8455F" },
  { v: "finestre", l: "Finestre di analisi", c: "#6C6C79" },
  { v: "chiusure", l: "Clienti chiusi", c: "#3FA45C" },
];

const TIPI_PERIODO = [
  { v: "test", l: "Test", d: "i primi tre giorni: si osserva, non si decide" },
  { v: "matura", l: "Matura", d: "i sette giorni successivi: la campagna si assesta" },
  { v: "attiva", l: "Attiva", d: "dal decimo giorno in poi, una settimana alla volta" },
];

/* Il budget giornaliero cambia nel tempo: per sapere quanto sarebbe dovuto uscire
   bisogna sapere quale importo era in vigore ogni singolo giorno. */
const budgetDelGiorno = (k, giorno) => {
  const st = (k.storicoBudget || []).filter((x) => x.data && x.data <= giorno)
    .sort((a, b) => (a.data < b.data ? -1 : 1));
  return st.length ? (st[st.length - 1].importo || 0) : (k.budgetGiorno || 0);
};

/* Il periodo davvero coperto dai dati: se ci sono fasi compilate vale il loro
   intervallo, altrimenti dalla partenza a oggi. Confrontare la spesa di una settimana
   con l'atteso di due mesi è il modo più rapido per generare allarmi falsi. */
const giorniTra = (a, b) => Math.round((new Date(b + "T00:00:00") - new Date(a + "T00:00:00")) / 864e5) + 1;

const coperturaDati = (data, k) => {
  const ps = (data.periodi || []).filter((pz) => pz.campagnaId === k.id
    && ((pz.spesa || 0) + (pz.impression || 0) + (pz.click || 0) + (pz.conversioni || 0) + (pz.ricavi || 0)) > 0);
  let dal = k.inizio || "", al = oggi(), daFasi = false;
  if (ps.length) {
    const d1 = ps.map((pz) => pz.inizio).filter(Boolean).sort()[0];
    const a1 = ps.map((pz) => pz.fine).filter(Boolean).sort().slice(-1)[0];
    if (d1) dal = d1;
    if (a1) al = a1;
    daFasi = true;
  }
  if (!dal) return { dal: "", al: "", daFasi, giorni: 0, sospetto: false };
  /* Una campagna spenta non accumula altri giorni: il conto si ferma alla sua fine */
  const limite = (k.stato !== "attiva" && k.fine && k.fine < oggi()) ? k.fine : oggi();
  if (al > limite) al = limite;
  if (al < dal) al = dal;
  const giorni = giorniTra(dal, al);
  /* Oltre un anno di finestra su una campagna attiva le date sono quasi certamente sbagliate */
  return { dal, al, daFasi, giorni, sospetto: giorni > 400 };
};

const erogazioneAttesa = (data, k) => {
  const conta = (dal, al) => {
    let t = 0, g = dal, guardia = 0;
    while (g <= al && guardia < 400) { t += budgetDelGiorno(k, g); g = sommaGiorni(g, 1); guardia++; }
    return t;
  };
  const ps = (data.periodi || []).filter((pz) => pz.campagnaId === k.id && pz.inizio
    && ((pz.spesa || 0) + (pz.impression || 0) + (pz.click || 0) + (pz.conversioni || 0) + (pz.ricavi || 0)) > 0);

  /* Con le finestre compilate conto solo i giorni davvero coperti: i buchi non contano.
     Una finestra più lunga di due mesi non è una finestra: è una data sbagliata. */
  if (ps.length) {
    let atteso = 0, giorni = 0, dal = null, al = null, storte = 0;
    ps.forEach((pz) => {
      const fine = !pz.fine || pz.fine > oggi() ? oggi() : pz.fine;
      if (fine < pz.inizio) { storte++; return; }
      const n = giorniTra(pz.inizio, fine);
      if (n > 60) { storte++; return; }
      atteso += conta(pz.inizio, fine); giorni += n;
      if (!dal || pz.inizio < dal) dal = pz.inizio;
      if (!al || fine > al) al = fine;
    });
    if (!giorni) return { atteso: 0, giorni: 0, daFasi: true, sospetto: true, storte };
    return { atteso, giorni, dal, al, daFasi: true, storte, sospetto: storte > 0 && giorni === 0 };
  }

  /* Senza finestre non so a quale periodo si riferisca la spesa che hai scritto:
     lo do per scontato solo su campagne partite da poco. */
  if (!k.inizio) return { atteso: 0, giorni: 0, daFasi: false, senzaDate: true };
  const eta = giorniTra(k.inizio, oggi());
  if (eta > 400) return { atteso: 0, giorni: eta, dal: k.inizio, al: oggi(), daFasi: false, sospetto: true };
  if (eta > 45) return { atteso: 0, giorni: eta, dal: k.inizio, al: oggi(), daFasi: false, senzaFinestre: true };
  return { atteso: conta(k.inizio, oggi()), giorni: eta, dal: k.inizio, al: oggi(), daFasi: false };
};



/* La data che Meta riporta non è sempre la partenza vera: viene riscritta quando
   una campagna viene riattivata. Il primo giorno in cui ha speso non mente. */
const inizioReale = (data, k) => {
  const prime = (data.periodi || [])
    .filter((p) => p.campagnaId === k.id && p.inizio && (p.spesa > 0 || p.impression > 0))
    .map((p) => p.inizio).sort();
  const daFasi = prime[0] || "";
  if (!k.inizio) return daFasi;
  if (!daFasi) return k.inizio;
  return daFasi < k.inizio ? daFasi : k.inizio;
};

const etaCampagna = (k) => {
  if (!k.inizio) return null;
  const fino = (k.stato !== "attiva" && k.fine && k.fine < oggi()) ? k.fine : oggi();
  const g = Math.floor((new Date(fino) - new Date(k.inizio)) / 864e5) + 1;
  return g < 0 ? 0 : g;
};

const faseCampagna = (k) => {
  const g = etaCampagna(k);
  if (g === null) return { g: null, fase: "nd", l: "senza data di inizio" };
  if (k.stato === "errore") {
    const gg = k.erroreDal ? giorniTra(k.erroreDal, oggi()) : 0;
    return { g, fase: "errore", l: gg ? `ferma da ${gg} ${gg === 1 ? "giorno" : "giorni"}` : k.causaErrore && k.causaErrore !== "pagamento" ? `ferma: ${causaFermo(k).l.toLowerCase()}` : "ferma per pagamento" };
  }
  if (k.stato !== "attiva" && k.fine && k.fine < oggi()) {
    return { g, fase: "chiusa", l: `${g} giorni · chiusa il ${dataIt(k.fine)}` };
  }
  if (g <= TEST_1) return { g, fase: "test", l: `Test · giorno ${g} di ${TEST_1}` };
  if (g <= TEST_2) return { g, fase: "matura", l: `Matura · giorno ${g} di ${TEST_2}` };
  if (g > 400) return { g, fase: "attiva", l: `data di partenza da controllare` };
  return { g, fase: "attiva", l: `Attiva · ${g} giorni` };
};

/* Andamento: differenza fra i totali di oggi e quelli dell'ultima rilevazione salvata.
   Il risultato è la resa del solo intervallo, non la media da inizio campagna. */
const ultimaRilevazione = (data, k) =>
  (data.rilevazioni || []).filter((r) => r.campagnaId === k.id).sort((a, b) => (a.data < b.data ? 1 : -1))[0];

const intervallo = (k, prec) => {
  if (!prec) return null;
  const sp = (k.spesa || 0) - (prec.spesa || 0);
  const cv = (k.conversioni || 0) - (prec.conversioni || 0);
  const cl = (k.click || 0) - (prec.click || 0);
  const imp = (k.impression || 0) - (prec.impression || 0);
  const rc = (k.ricavi || 0) - (prec.ricavi || 0);
  if (sp <= 0) return null;
  /* Lead e clic sul link: solo se la rilevazione precedente li aveva salvati,
     altrimenti la differenza sarebbe il totale da inizio campagna. */
  const leadDi = (x) => x.lead || (k.obiettivo === "lead" ? x.conversioni || 0 : 0);
  const ld = prec.lead !== undefined || k.obiettivo === "lead" ? Math.max(0, leadDi(k) - leadDi(prec)) : 0;
  const clk = prec.clickLink !== undefined ? Math.max(0, (k.clickLink || 0) - (prec.clickLink || 0)) : 0;
  return {
    da: prec.data, sp, cv, cl, imp, rc, ld, clk,
    cpa: cv > 0 ? sp / cv : 0, roas: sp ? rc / sp : 0, ctr: imp ? (cl / imp) * 100 : 0,
    ctrLink: imp ? ((clk || cl) / imp) * 100 : 0, cpl: ld ? sp / ld : 0, soloTotali: clk === 0 && cl > 0,
  };
};

/* ================================================================== */
/* Guida                                                               */
/* ================================================================== */

const GUIDA = [
  {
    id: "inizio", t: "Primi passi",
    intro: "L'ordine con cui riempi le cose conta: ogni sezione si appoggia alla precedente. Segui questa sequenza e non ti ritrovi con schermate vuote.",
    blocchi: [
      { h: "1. Identità dell'agenzia", p: "Impostazioni → Identità: nome, sottotitolo, logo e soprattutto il costo orario del tuo tempo, che serve al conto economico per cliente." },
      { h: "2. I clienti", p: "Clienti → Aggiungi. Oltre ai dati anagrafici, compila margine lordo e resi: senza quei due numeri il ROAS resta un dato senza significato economico. Nella stessa scheda ci sono i dati di fatturazione — ragione sociale, partita IVA, sede legale, codice destinatario o PEC, referente amministrativo, termini e metodo di pagamento — che poi ritrovi raccolti nella Scheda cliente con un pulsante per copiarli quando devi emettere fattura." },
      { h: "3. Le campagne", p: "Entra nel cliente, sezione Advertising. Puoi inserirle a mano o importarle da file: in quel caso arrivano già con date, stato e budget giornaliero. Due cose vanno impostate comunque e sono quelle su cui si regge il resto: il CPA o CPL obiettivo, che è il KPI del semaforo, e il budget giornaliero vero della piattaforma, che serve a misurare l'erogazione." },
      { h: "4. Le finestre", p: "Dal pannello Fasi genera le finestre di analisi: tre giorni di Test, sette di Matura, poi settimane. Senza, l'app giudica sulla media dall'inizio, che su una campagna lunga non dice come va adesso." },
      { h: "5. Il commerciale", p: "Ciclo di vendita per i numeri del mese, Chiusure per ogni cliente chiuso. Le chiusure hanno la precedenza: quando ci sono, riempiono da sole i totali del mese." },
      { h: "6. Il backup", p: "Backup → Scarica backup. Fallo subito la prima volta, così sai dove finisce il file e come si rimette a posto." },
    ],
  },
  {
    id: "routine", t: "La routine",
    intro: "Lo strumento non serve a niente se i dati invecchiano. Questa è la cadenza minima perché i numeri restino affidabili.",
    blocchi: [
      { h: "La panoramica", p: "In cima ci sono sei indicatori del mese in corso — spesa, lead, CPL, costo per cliente acquisito, contratti firmati e ritorno — ciascuno con la variazione rispetto al mese precedente. Sulle metriche di costo la freccia in giù è verde: scendere è un miglioramento. Sotto, l'andamento di contratti e spesa negli ultimi dodici mesi e la distribuzione della spesa per piattaforma." },
      { h: "Cosa entra nei totali dell'agenzia", p: "La panoramica e il riepilogo mensile a livello agenzia contano solo i clienti principali: spesa, lead, CPL, contratti e andamento escludono i sottoclienti, e nella tabella Performance la riga del cliente principale ha solo i suoi numeri: sotto il nome leggi quanti sottoclienti sono esclusi, e la freccia li apre uno per riga. Per i numeri del gruppo sommati entri nel cliente madre e usi l'interruttore consolidato." },
      { h: "L'eccezione degli avvisi", p: "Gli avvisi restano su tutti gli account, sottoclienti compresi, e nella riga del cliente principale è indicato quanti arrivano da loro. È una scelta deliberata: un problema critico sulla campagna di un sottocliente non deve restare invisibile solo perché i suoi numeri non entrano nei totali." },
      { h: "Sforzo/Resa", p: "Quanto incassi per ogni tentativo di contatto: incassato diviso chiamate fatte, nel mese. È il valore medio che ogni contatto gestito porta davvero in cassa, e si legge accanto al CPL. Se è più alto del CPL il funnel è in profitto e aumentare il volume ha senso; se è più basso, ogni contatto in più costa più di quanto rende, e la colonna diventa bordeaux. Scomposto, quel numero è tasso di conversione per scontrino medio: tanti contatti con poco incassato vuol dire lead fuori target o vendita inefficace; pochi contatti con molto incassato vuol dire un funnel di qualità da scalare tenendo fermo il target. Se nel ciclo di vendita non hai registrato le chiamate, la colonna usa i contatti entrati nel funnel e lo dichiara sotto il numero: le due basi non sono confrontabili fra loro, quindi tienine conto quando metti due clienti a fianco." },
      { h: "Performance dei clienti", p: "La tabella riporta per ogni cliente la spesa, i lead e il CPL dell'ultimo mese con dati, e le demo, le chiusure, i contratti e lo Sforzo/Resa dell'ultimo mese che ha numeri commerciali: spesso il mese in corso ha già la spesa di Meta ma non ancora il ciclo di vendita, e in quel caso sotto il numero compare il mese da cui arriva. Più la resa rispetto all'obiettivo pubblicitario. Il pallino dello stato riassume il giudizio sui risultati: rosso (Critico) se ha avvisi critici sulle campagne, rosso con la scritta Campagne ferme se l'unico problema è un fermo per pagamento o account, ambra se ha avvisi di attenzione o è sotto obiettivo, verde se regge. I task in ritardo restano negli Avvisi ma non decidono lo stato. Sotto lo stato compare la lettura del Ciclo di vendita: «in attivo» con quante volte i contratti superano la spesa, oppure il rapporto e il pareggio quando il cliente non lo supera. Cliccando entri nel cliente." },
      { h: "Ogni giorno, due minuti", p: "Apri il Cruscotto e guarda gli avvisi. Se ce n'è uno critico, o lo risolvi o lo trasformi in un task con una data." },
      { h: "Ogni settimana, venti minuti", p: "Aggiorna i dati delle campagne, poi premi Salva rilevazione in Advertising. È quel gesto che crea lo storico: senza, l'app può dirti solo se sei sopra soglia, non se stai peggiorando." },
      { h: "Ogni mese, un'ora", p: "Compila il periodo del Ciclo di vendita, registra le chiusure, genera il Report del cliente e scarica un backup. Poi guarda Economia: è il momento in cui scopri quale cliente ti sta costando più di quanto rende." },
      { h: "Quando cambi qualcosa", p: "Scrivi una riga nel Diario. Trenta secondi adesso valgono un'ora di ricostruzioni fra due mesi, quando guarderai una curva storta senza ricordare cosa avevi toccato." },
      { h: "Se la routine ti sfugge", p: "In File ci sono tre procedure pronte — rilevazione settimanale, chiusura del mese, nuovo cliente — con i passaggi in ordine. E se non sai da dove partire, il Consulente legge i tuoi numeri e ti dice su cosa lavorare per primo." },
    ],
  },
  {
    id: "struttura", t: "Agenzia e sotto-account",
    intro: "In alto a sinistra c'è il selettore account: è la cosa più importante da capire dell'interfaccia.",
    blocchi: [
      { h: "Livello agenzia", p: "Con \"Tutti i clienti\" selezionato vedi il portafoglio intero: panoramica, semaforo, advertising di tutti, economia, automazioni, importazione e backup." },
      { h: "Sotto-account", p: "Scegli un cliente ed entri nel suo ambiente: contatti, pipeline, campagne, funnel, ciclo di vendita, chiusure, report. Tutto quello che crei da dentro nasce già assegnato a quel cliente." },
      { h: "Menu a icone", p: "La freccia « accanto al logo riduce il menu a una colonna di icone. Su schermi stretti parte già ridotto; il pallino bordeaux su un'icona segnala avvisi aperti." },
      { h: "Il selettore nella barra", p: "Nella colonna di sinistra, sopra le sezioni, c'è il menu Account: scegli un cliente per entrare nel suo account, o “Tutti i clienti” per tornare alla vista agenzia. I sottoclienti compaiono rientrati sotto il loro cliente principale. Con il menu ridotto a icone resta il logo dell'account aperto: cliccandolo il menu si riapre." },
      { h: "In gestione dal", p: "Nella scheda del cliente, sotto la sede, c'è la data da cui lo segui. All'inizio è quella in cui hai creato la scheda, ma puoi cambiarla: se un avvocato è entrato nel portafoglio a marzo e l'hai inserito a settembre, metti marzo. La colonna Data partenza negli elenchi mostra invece il primo giorno di attività pubblicitaria, e ripiega su questa data quando non ci sono ancora campagne. È anche importabile da foglio, con una colonna \"In gestione dal\", \"Onboarding\" o \"Cliente dal\"." },
      { h: "Data contratto", p: "Accanto c’è la data di firma, che è un’altra cosa: il contratto può essere firmato prima dell’avvio o dopo un periodo di prova. Quando le due date differiscono, la scheda ti dice di quanti giorni e in che verso. Anche questa si importa, con una colonna \"Data contratto\", \"Firma\" o \"Firmato il\"." },
      { h: "Due elenchi distinti", p: "La sezione Clienti conta e mostra solo i clienti principali: se ne segui tre, il menu dice tre, anche quando sotto di loro ci sono ottanta account. La freccia accanto al nome apre i suoi sottoclienti nella stessa tabella; cercando un nome compaiono già aperti. I sottoclienti hanno anche la loro sezione, con un riquadro chiuso per ogni cliente principale: cliccandolo vedi spesa, lead, CPL, resa e avvisi di ciascuno e i pulsanti per archiviarli o eliminarli. In cima trovi i totali del gruppo, che restano esclusi dai totali dell'agenzia." },
      { h: "La freccia dei sottoclienti", p: "Ovunque una tabella elenca più clienti si vedono a colpo d'occhio i clienti principali, e la freccia ▸ accanto al nome apre i loro sottoclienti: in Clienti, Advertising, Ciclo di vendita, Chiusure, Economia, nella tabella Performance della Panoramica e nel collegamento con Meta in Impostazioni. Dove la tabella mostra totali (Advertising, Ciclo di vendita, Chiusure, Economia) la riga chiusa somma tutto il gruppo e, aperta, mostra il cliente principale da solo e ogni sottocliente. In Performance la riga del principale resta con i suoi numeri, come nei totali dell'agenzia, e aperta mostra i numeri di ogni sottocliente. In Impostazioni, sotto il nome, leggi quanti sottoclienti hanno già un account collegato. Ogni sezione ricorda le sue frecce aperte finché non ricarichi la pagina." },
      { h: "Le campagne nel cruscotto", p: "La scheda Campagne del cruscotto cliente parte dalla più recente e ha i suoi filtri: ordinamento, stato, fase e segnale del semaforo. Sono indipendenti da quelli della sezione Advertising, così puoi tenere il cruscotto su \"solo i rossi\" senza cambiare come guardi l'altra schermata. Anche qui ogni riga ha il pallino del semaforo all'inizio." },
      { h: "Sottoclienti", p: "Un cliente può contenerne altri. Li crei da tre punti: la scheda Sottoclienti nel cruscotto del cliente, il pulsante in fondo alla riga nell’elenco Clienti, oppure il campo \"Sottocliente di\" nel modulo, che serve anche a spostare sotto un altro un cliente già esistente. Serve quando gestisci l'azienda madre e insieme i suoi clienti finali — per esempio una società di consulenza e i singoli professionisti che entrano nel suo portafoglio. Ogni sottocliente ha la struttura completa: campagne, funnel, messaggio, ciclo di vendita, chiusure, task e file suoi." },
      { h: "Ordinati per cognome", p: "I sottoclienti compaiono sempre in ordine di cognome, non di nome: l'app ignora i titoli come Avv. o Studio Legale, tiene attaccate le particelle — De Luca sotto la D, Di Marco sotto la D — e scarta le forme societarie, così \"Studio Legale Rossi & Partners\" si ordina su Rossi. Se un nome è scritto al contrario — cognome prima, come in “Laudando Antonio” — la deduzione sbaglia: in quel caso compila il campo “Cognome per l\u2019ordinamento” nella scheda del cliente, che vince sempre sulla deduzione. Si importa anche da foglio, con una colonna Cognome." },
      { h: "Consolidato o singolo", p: "Quando entri in un cliente che ha sottoclienti, in alto a destra compare l'interruttore: consolidato somma tutto il gruppo, l'altra modalità mostra solo l'account madre. Vale per ogni sezione, dal cruscotto ad Advertising. Nel cruscotto trovi anche la tabella dei sottoclienti con spesa, resa e avvisi di ciascuno." },
      { h: "Archiviare un cliente", p: "Quando smetti di seguirlo ma potrebbe tornare, archivialo invece di eliminarlo: dal pulsante in fondo alla riga nell'elenco Clienti o dalla sua Scheda. Esce dall'elenco, dal selettore, dai totali, dal semaforo, dai grafici, dagli avvisi e dal contesto del consulente. I sottoclienti lo seguono." },
      { h: "Dove finisce lo storico", p: "Tutto quello che riguarda quel cliente — campagne, fasi, funnel, contatti, trattative, periodi commerciali, chiusure, task, file, automazioni — viene spostato in un deposito. Non è cancellato e non è più nei conti correnti: è messo da parte. Nella scheda Clienti archiviati vedi per ciascuno quante voci ci sono, quanta spesa storica e quanti contratti." },
      { h: "Riattivare", p: "Il pulsante Riattiva chiede sempre come vuoi ripartire: con lo storico, che riporta tutto dentro, oppure da zero, che riattiva il cliente lasciando i dati nel deposito. Nel secondo caso, nella sua Scheda compare il riquadro \"Storico in deposito\" con il pulsante per reinserirlo in qualsiasi momento — utile se il rapporto riparte davvero e i numeri vecchi tornano a servire." },
      { h: "Eliminare davvero", p: "L'eliminazione resta per i casi in cui i dati non servono più a nessuno. I pulsanti sono tre: in fondo alla riga nell'elenco Clienti, nella tabella Sottoclienti del cruscotto — così cancelli un singolo sottocliente senza entrarci — e nella Scheda cliente. In tutti i casi c'è il pulsante: prima di procedere ti elenca esattamente cosa sparisce — sottoclienti compresi, con tutto quello che contengono — e accanto trovi il pulsante per scaricare un backup. Se invece il rapporto è finito ma i dati ti servono ancora, metti lo stato su \"Chiuso\": esce dal semaforo e dai conteggi attivi, ma resta consultabile." },
    ],
  },
  {
    id: "adv", t: "Advertising e avvisi",
    intro: "Inserisci solo i dati grezzi che copi dalla piattaforma. CTR, CPC, CPM, CPA e ROAS li calcola l'app.",
    blocchi: [
      { h: "Clienti in attivo: niente allarmi critici", p: "Gli avvisi tengono conto del Ciclo di vendita. Se negli ultimi tre mesi con numeri commerciali i contratti del cliente superano la spesa, oltre il punto di pareggio quando nella scheda è impostato il margine, altrimenti oltre 1×, il cliente è in attivo: CPL sopra il tetto, CTR basso, CPM alto e simili restano segnalati, ma come attenzione e con il motivo scritto nell'avviso, non come critici. Critici restano solo i fermi, cioè errore di pagamento e account disattivato. Un sottocliente senza ciclo di vendita suo segue il giudizio del cliente madre. Gli avvisi si ricalcolano a ogni apertura, quindi vale anche per quelli già comparsi." },
      { h: "Campagne a lead", p: "Nelle campagne con obiettivo Lead il risultato è il lead: l'avviso di spesa senza risultati scatta solo se mancano i lead, non le conversioni d'acquisto, che Meta su queste campagne non registra." },
      { h: "Più account pubblicitari per lo stesso cliente", p: "Un cliente può avere quanti account vuoi: li colleghi dalla sua scheda, sezione Meta Ads. Dentro il cliente, Advertising mostra allora la tabella Per obiettivo e account: per ogni obiettivo una riga di totale, con spesa, lead, CPL, conversioni, CPA, ROAS e contratti sommati su tutti gli account, e sotto una riga per ogni account con i suoi numeri. Obiettivi diversi restano separati, perché un CPL di lead e un CPA di vendite non si sommano. Un clic su un account filtra l'elenco delle campagne su quell'account e quell'obiettivo; nei filtri c'è anche la scelta Account pubblicitario, e sotto ogni campagna compare il nome dell'account da cui arriva." },
      { h: "Clienti principali e sottoclienti", p: "In Advertising si vedono a colpo d'occhio solo i clienti principali, con i numeri loro e dei sottoclienti sommati. La freccia accanto al nome apre i sottoclienti, uno per riga, più la riga dell'account principale da solo. Dentro un cliente vale lo stesso: prima le sue campagne, poi una riga chiusa per ogni sottocliente con i suoi totali, che la freccia apre sulle sue campagne." },
      { h: "Dentro e fuori dal cliente", p: "Con Tutti i clienti selezionato, Advertising mostra una riga per cliente con i totali: campagne e quante sono accese, spesa, lead, CPL, conversioni, ROAS, contratti e incassato, ognuno nella valuta del suo cliente. Contratti e incassato arrivano dalle schede chiusura collegate alle campagne o dagli importi scritti nelle campagne; se non ce ne sono, dal Ciclo di vendita del cliente (sotto il numero è indicato quanti mesi), e in quel caso il ROAS, quando la piattaforma non dichiara ricavi, è il rapporto fra contratti e spesa di quei soli mesi. Campagne, ad set, annunci, finestre, ritmo di erogazione e totali per funnel e linea si guardano dentro il cliente: un clic sulla riga ti porta lì." },
      { h: "Benchmark", p: "Per ogni campagna imposti ROAS minimo, CPL massimo, CPA massimo, CTR minimo, frequenza e CPM massimi. Lascia 0 quello che non vuoi controllare. Cambiando obiettivo l'app propone soglie tipiche, che poi correggi con i numeri veri del cliente." },
      { h: "CPL e CPA sono due cose diverse", p: "Nelle campagne a contatti il lead è il primo punto di conversione, e il CPL massimo è il tetto che conta: è lì che la campagna si rompe per prima. Il CPA massimo misura il punto dopo — l'appuntamento, la vendita, qualunque cosa tu stia contando come conversione. Se li tieni separati vedi dove si perde il margine: un CPL in linea con un CPA fuori tetto è un problema di vendita, non di advertising. Se lasci il CPL a 0 vale il CPA, così le campagne create prima di questo campo continuano a essere giudicate; e dove lead e conversioni sono lo stesso numero l'app mostra solo l'avviso sul CPL, per non dire due volte la stessa cosa." },
      { h: "Il lead non è una conversione", p: "Sono due colonne diverse apposta. Il lead è il contatto che arriva; la conversione è quello che succede dopo — l'appuntamento fissato, il preventivo accettato, la vendita — e quel numero la piattaforma non lo conosce: lo sai tu. Meta però chiama \"risultati\" i lead, quindi importando un account a contatti finiscono in tutte e due le colonne e il CPA esce identico al CPL. Quando capita, sotto le conversioni compare la scritta \"= lead\" e in cima alla sezione appare il pulsante \"Separa conversioni dai lead\": azzera le conversioni sulle campagne che stai guardando e sulle loro finestre, lasciando intatti lead e CPL. Da lì il campo Conversioni torna tuo: lo scrivi nella scheda anche se la campagna ha le finestre — è l'unico dato del blocco che resta scrivibile, insieme ai ricavi — e accanto al numero compare \"a mano\" perché si veda che non arriva da Meta." },
      { h: "Come leggere un avviso", p: "Ogni avviso dice tre cose separate: cosa succede con i numeri, la causa probabile, e la mossa consigliata. Sono scritti per evitare l'errore più comune, cioè reagire alla metrica sbagliata: davanti a un CPA alto ti fanno confrontare CPC e tasso di conversione, perché quasi sempre il problema è a valle del click." },
      { h: "Ritmo di erogazione", p: "La barra mostra quanto budget hai speso, la tacca verticale dove sei nel periodo. Se il riempimento supera la tacca stai bruciando il budget in anticipo; se resta molto indietro rischi di chiudere il mese sotto erogazione, che con un cliente a budget è un problema di rapporto oltre che di risultato." },
      { h: "Ad set e annunci", p: "È il livello a cui ottimizzi davvero. Scrivi i nomi esattamente come nella piattaforma e come nelle schede chiusura: la colonna Contratti si aggancia per corrispondenza esatta del nome." },
      { h: "Campagne, funnel e totali", p: "Ogni campagna può essere collegata a un funnel: da lì nascono la colonna Funnel, il filtro dedicato e la tabella Totali per funnel, che ti dice quanto costa alimentare ciascun percorso. In fondo alla tabella delle campagne c'è la riga \"Tutte insieme\": è la somma come se fosse una campagna unica." },
      { h: "Perché le medie non si mediano", p: "CTR, CPC, CPA e ROAS della riga totale non sono la media delle colonne: sono ricalcolati dai totali — click su impression, spesa su conversioni, ricavi su spesa. Fare la media delle medie farebbe pesare uguale la campagna da cento euro e quella da diecimila, ed è l'errore che rende irriconoscibili metà dei report che girano." },
      { h: "Le finestre di valutazione", p: "Una campagna non si giudica sempre allo stesso modo. Nei primi tre giorni l'app non emette nessun giudizio: compare solo una nota che dice di annotare e non toccare, perché la consegna è ancora instabile. Dal quarto al decimo giorno è la fase Matura: gli avvisi ci sono ma non diventano mai critici, tranne quelli tecnici — tracciamento, erogazione, margine — che critici lo sono a prescindere. Dall'undicesimo giorno la campagna è Attiva e viene giudicata sulla finestra degli ultimi sette giorni, non sulla media da inizio campagna." },
      { h: "Perché la finestra cambia tutto", p: "Una campagna che gira da due mesi ha una media che si muove pianissimo: può peggiorare per tre settimane senza che il totale se ne accorga. Valutando gli ultimi sette giorni, gli avvisi parlano di adesso. La finestra si costruisce dalle rilevazioni salvate: senza rilevazioni l'app torna al totale, ed è un'altra ragione per fare quel salvataggio ogni settimana." },
      { h: "Budget giornaliero", p: "Il budget della campagna si indica al giorno, come lo imposti in piattaforma. Da lì l'app calcola quanto avrebbe dovuto uscire dall'inizio a oggi e lo confronta con la spesa reale: se stai erogando oltre il 20% in più ti avvisa, se stai sotto il 25% pure. Il budget totale del periodo resta un campo facoltativo." },
      { h: "Tre numeri diversi per i soldi", p: "Ricavi sono quelli che dichiara la piattaforma, Contratto è quanto è stato firmato, Incassato quanto è entrato davvero. Contratto e incassato li scrivi direttamente nella scheda della campagna o della singola finestra; se li lasci a zero, l’app li ricava dalle schede della sezione Chiusure collegate a quella campagna, e quando li scrivi a mano vincono quelli. Quando guardi una singola finestra, contratto e incassato sono filtrati per data di chiusura." },
      { h: "Lead e CPL dappertutto", p: "Ogni riepilogo che mostra la spesa mostra anche i lead e il costo per lead: tabella campagne, ad set e annunci, totali per funnel, per linea e per piattaforma, archivio, elenco clienti, sottoclienti e riepilogo mensile. Il CPL diventa bordeaux quando supera il CPL obiettivo della campagna, o il CPA se il CPL non è impostato. Dove il lead non è un dato di campagna ma di periodo — elenco clienti, sottoclienti, panoramica e riepilogo mensile — il valore è quello dell'ultimo mese con dati, non del totale storico: sommare i lead di due anni accanto alla spesa complessiva darebbe un CPL che non corrisponde a nessun periodo reale. In quei punti i lead del mese si prendono nell'ordine da chi li conosce meglio: prima le giornate dei venditori, poi il ciclo di vendita, e se nessuno dei due è compilato le finestre pubblicitarie di quel mese. Serve per gli account che vivono di lead importati da Meta e non tengono un ciclo di vendita: prima quelle colonne restavano vuote." },
      { h: "Clic totali e clic sul link", p: "Sono due numeri diversi e le piattaforme li mostrano entrambi: i clic totali contano anche like, commenti ed espansioni del testo, quelli sul link contano chi è davvero uscito verso la tua pagina. CTR link e CPC link si calcolano su questi ultimi, ed è su quelli che va letta la soglia dello 0,80%. Se non compili i clic sul link l'app ripiega sui totali e te lo scrive sotto il numero, perché in quel caso il CTR risulta più alto del vero." },
      { h: "Lead e costo per lead", p: "Nelle campagne a lead il CPL è la metrica che porti al cliente. Compila il campo Lead — anche quando coincide con le conversioni — e trovi il CPL in tabella, nel pannello delle fasi e nella riga dei totali, calcolato come spesa divisa per lead e non come media dei CPL." },
      { h: "Cambiare il budget e tenerne traccia", p: "Nella scheda della campagna, sotto il giornaliero, c'è lo storico delle variazioni: ogni riga ha data, importo e una nota sul perché. Ne aggiungi quante ne vuoi con \"+ Nuovo budget giornaliero\", e ciascuna vale dal giorno indicato fino alla successiva. Il giornaliero attuale si allinea da solo all'ultima variazione." },
      { h: "Perché serve davvero", p: "L'erogazione attesa non è budget per giorni: l'app somma giorno per giorno l'importo in vigore in quel momento, e solo sui giorni davvero coperti dai dati che hai inserito — se hai compilato due settimane, confronta due settimane. Se il 10 hai alzato da 50 a 120 euro, il confronto con lo speso resta corretto — senza storico ti direbbe che stai spendendo il doppio del previsto, e prenderesti una decisione sbagliata su un dato sbagliato." },
      { h: "Il pannello dei filtri", p: "In Advertising i filtri stanno dietro un solo pulsante, che mostra quanti ne hai attivi; accanto, in chiaro, cosa stai guardando — per esempio \"12 campagne · dalla più recente\". Aprendolo trovi in griglia cliente, ordinamento, fase corrente, semaforo, stato, numeri da mostrare, piattaforma, obiettivo, funnel e linea di business. Il filtro semaforo isola le campagne per segnale — per esempio solo quelle da spegnere — e il segnale scelto compare accanto al conteggio con il suo pallino. Sotto, una striscia con i totali di quello che hai selezionato, così vedi l'effetto del filtro senza scorrere la tabella." },
      { h: "Ordinare e filtrare", p: "L'elenco parte dalla campagna più recente, che è quella su cui devi decidere oggi. Dal primo menu cambi ordine: per spesa, per lead, per CPL dal più basso, per ROAS dal più alto, per nome. Accanto c'è il filtro per fase corrente — Test, Matura, Attiva, chiuse, ferme per pagamento, senza data di inizio — che è cosa diversa dal menu della fase da visualizzare: quello sceglie quali numeri leggere, questo quali campagne mostrare. Il pulsante Azzera riporta tutto, ordinamento compreso." },
      { h: "Come si misura l'erogazione", p: "Il confronto è fra due numeri omogenei: quanto la campagna ha speso in media in un giorno, calcolato sulla spesa reale divisa per i giorni coperti dalle finestre, e quanto dovrebbe spenderne secondo il budget giornaliero impostato in piattaforma. Nessuna proiezione e nessun totale, perché un totale previsto su un periodo lungo gonfia qualunque scarto e rende il numero illeggibile." },
      { h: "Il budget non si inventa", p: "Se nella scheda non c'è un budget giornaliero, l'app non lo deduce dividendo un budget totale per la durata: lascia il campo vuoto e non calcola l'erogazione. Un numero inventato produrrebbe confronti falsi, e un confronto falso è peggio di un confronto assente." },
      { h: "Quando il budget in scheda è sbagliato", p: "Se la spesa reale è sotto un decimo del budget indicato, l'app non parla di sotto-erogazione: segnala che il numero in scheda non è credibile e dice di quante volte è fuori scala. Di solito è un budget totale finito nel campo del giornaliero. Nel semaforo, in quel caso, il ritmo di spesa non viene giudicato affatto." },
      { h: "La data di partenza vera", p: "La data che la piattaforma riporta non è sempre quella di avvio: viene riscritta quando una campagna viene riattivata, o impostata al momento della creazione anche se la consegna è iniziata prima. Il CRM confronta la data in scheda con la prima finestra in cui risulta spesa e usa la più antica delle due: fase, età, erogazione e semaforo partono da lì. Quando le due date differiscono, accanto al nome della campagna compare un'etichetta con la data usata, e passandoci sopra leggi quella in scheda." },
      { h: "Le tre fasi", p: "Test sono i primi tre giorni, Matura i sette successivi, Attiva dal decimo giorno in poi, una settimana alla volta. La fase corrente compare in tabella accanto allo stato, e cambia il modo in cui l'app giudica: in Test nessun giudizio, in Matura nessun avviso critico salvo quelli tecnici, in Attiva la valutazione si fa sulla finestra dei sette giorni." },
      { h: "Le finestre si impostano nella scheda", p: "Nel modulo della campagna ci sono la data di partenza, la data di fine — da compilare solo quando la spegni o la concludi; se resta vuota vale la fine dell'ultima finestra — e sotto il blocco delle fasi. I pulsanti aggiungono la fase Test e la fase Matura con le date già calcolate, e \"+ Settimana successiva\" attacca ogni volta la finestra seguente al giorno dopo la fine della precedente. Le date restano modificabili a mano quando la realtà non segue il calendario." },
      { h: "Ad set e annunci in ordine", p: "La tabella segue la gerarchia reale: prima l'ad set, poi i suoi annunci rientrati sotto di lui, campagna per campagna. Nella scheda dell'annuncio c'è il campo \"Ad set di appartenenza\", che elenca gli ad set della campagna scelta: è quel collegamento a costruire l'ordine. Gli annunci non ancora collegati restano a sinistra con la dicitura \"da collegare a un ad set\", perché non sono figli di nessuno e non devono sembrarlo." },
      { h: "Importarli già collegati", p: "Nel file degli ad set e annunci puoi includere due colonne in più: il livello e il nome dell'ad set di appartenenza. L'app deduce da sole se una riga è un ad set o un annuncio — se indichi un ad set padre diverso dal nome della riga, è un annuncio — e lo aggancia. Vale per i clienti e per i sottoclienti allo stesso modo." },
      { h: "I numeri di ogni finestra", p: "Dal pulsante Fasi su ogni riga apri l'archivio delle finestre, ciascuna con le sue date e i suoi numeri. Il pulsante \"Genera le finestre\" le prepara vuote a partire dalla data di inizio, con i tagli giusti: tu inserisci solo i dati. Attenzione a un punto: i numeri di ogni finestra sono quelli del solo periodo, non cumulati dall'inizio." },
      { h: "Caricare un file", p: "Il campo in cima alla sezione accetta Excel, CSV e TSV, e puoi anche trascinare il file dentro il riquadro tratteggiato. Sotto, un suggerimento cambia con il tipo scelto e ti dice cosa si aspetta. Lo stesso vale per il registro dei venditori nel Ciclo di vendita." },
      { h: "Importare le campagne complete", p: "Il file delle campagne non porta solo i numeri: riconosce anche data di partenza, data di fine, stato, piattaforma, obiettivo e budget giornaliero. Gli stati scritti come li esporta Meta — ACTIVE, PAUSED, COMPLETED — vengono tradotti nei tuoi, e quelli con problemi di pagamento diventano \"Errore pagamento\". Senza queste colonne le campagne nascono tutte con la data di oggi e in stato attivo, e il filtro per fase non ha niente da distinguere." },
      { h: "Importare le fasi", p: "Il tipo \"Fasi delle campagne\" carica un file con una riga per campagna e per settimana: nome campagna, fase, date e numeri del periodo. Le finestre si agganciano alla campagna per nome esatto, e stessa campagna con stessa data di inizio significa stessa finestra — viene aggiornata, non duplicata. È il formato prodotto dallo script di estrazione da Meta." },
      { h: "Guardare una fase alla volta", p: "Il primo menu sopra la tabella cambia cosa stai leggendo: totale campagna, solo fase Test, solo fase Matura, ultima settimana attiva, oppure tutte le settimane attive sommate. Le campagne che non hanno quella finestra spariscono dall'elenco invece di mostrare zeri. Quando una campagna ha le fasi registrate, il suo totale non è più il numero scritto a mano nella scheda: è la somma delle finestre." },
      { h: "Errore di pagamento", p: "Fra gli stati della campagna c'è \"Errore pagamento\": selezionalo appena te ne accorgi. Da quel momento compare un avviso critico che resta finché non cambi stato, e conta i giorni di fermo — perché il danno di un errore di pagamento non è l'errore, ma i giorni in cui nessuno se ne accorge. La campagna esce dai conteggi delle attive, così non falsa medie e confronti." },
      { h: "Quando lo risolvi", p: "Riporta lo stato su Attiva: l'avviso sparisce e l'app scrive da sola nel Diario quanto è durato il fermo, con le date di inizio e fine. Resta lo storico, che serve quando un cliente chiede perché quel mese ha speso meno, o quando lo stesso metodo di pagamento salta per la terza volta." },
      { h: "Chiudere una campagna", p: "Quando spegni o concludi una campagna, compila la Data di fine nella sua scheda. Non è un dettaglio anagrafico: da quel giorno l'app smette di contare i giorni e di calcolare l'erogazione attesa, così una campagna ferma da tre mesi non continua ad accumulare budget previsto che nessuno ha speso. In tabella, nella colonna Fase, leggi la durata e il giorno di chiusura. Se lo stato è concluso o in pausa ma la data manca, compare una nota che te lo ricorda." },
      { h: "Quando le date non tornano", p: "Se una campagna risulta partita da più di quattrocento giorni, o se una finestra dura più di due mesi, l'app non calcola l'erogazione: te lo dice e basta, perché quasi sempre è un anno sbagliato nella data di partenza. Le finestre con date impossibili vengono escluse dal conteggio invece di gonfiare l'atteso. Se succede, correggi la data in scheda, cancella le finestre e usa \"Genera le finestre\"." },
      { h: "Quando non può valutare il ritmo", p: "Su una campagna che gira da più di quarantacinque giorni con un totale unico scritto a mano, l'app non sa a quali giorni si riferisca quella spesa: invece di inventare un confronto ti chiede di generare le finestre. Nessun avviso viene emesso su campagne senza spesa registrata." },
      { h: "Archiviare tutto l'advertising", p: "Dal riquadro Archivio, dentro un cliente, il pulsante \"Archivia tutto l'advertising\" sposta in deposito campagne, ad set, finestre e rilevazioni in un colpo solo. Funnel, ciclo di vendita, chiusure, task e file restano intatti: serve quando la parte pubblicitaria riparte da capo ma il rapporto continua. Se il cliente ha sottoclienti, l'app chiede se includerli o fermarsi all'account madre." },
      { h: "Riattivare l'advertising", p: "Ogni deposito compare nello stesso riquadro con la data, il numero di campagne e finestre e la spesa storica. Il pulsante Riattiva chiede come ripartire: con lo storico, e tutto torna nei conti, nelle medie e negli avvisi; oppure da zero, e i dati restano in deposito così le campagne nuove non vengono confrontate con le vecchie. La scelta non è definitiva: il deposito resta lì finché non lo reinserisci." },
      { h: "Archiviare le vecchie campagne", p: "Una campagna archiviata esce da tutto: medie, totali, semaforo, grafici, avvisi e contesto del consulente. I dati non si perdono, restano nel riquadro Archivio in fondo alla sezione, con i loro totali separati e il pulsante per ripristinarli. C'è anche un comando per archiviare in blocco tutte quelle in stato \"conclusa\": è la pulizia da fare a fine mese, perché una campagna spenta sei mesi fa continuerebbe a schiacciare verso il basso il CTR medio e a falsare il costo per acquisizione." },
    ],
  },
  {
    id: "messaggio", t: "Messaggio e Core 4",
    intro: "La sezione Messaggio di ogni cliente raccoglie le risposte da cui nasce tutto il resto, e da quelle genera quattro annunci già scritti, uno per fase del percorso d'acquisto.",
    blocchi: [
      { h: "La frase che tiene insieme tutto", p: "Compilando mercato, desiderio e tensione emotiva, l'app compone la frase del tipo \"Aiuto [mercato] a [desiderio], senza [tensione]\". Se quella frase non sta in piedi, nessun annuncio la salverà: è il primo controllo da fare su un cliente nuovo." },
      { h: "Chi compra e chi guarda soltanto", p: "Due campi separati per i dolori di chi compra e i dolori dei curiosi. Sono entrambi veri, ma solo i primi sono attuali e urgenti: costruire il messaggio sui secondi riempie il calendario di persone che non comprano, e poi si dà la colpa al commerciale." },
      { h: "Il livello di consapevolezza", p: "Da chi cerca già la soluzione a chi non ha ancora nominato il problema. L'app ti suggerisce il tipo di gancio adatto: dichiarare il risultato, alzare la promessa, presentare un meccanismo unico, presentarne uno aggiornato, oppure partire dal sintomo quotidiano." },
      { h: "I quattro annunci", p: "Uno per fase: freddo con problema, soluzione e benefici; tiepido con la guida all'acquisto; tiepido con il chi siamo che supera l'obiezione più grande; caldo con l'offerta e la proiezione nel futuro. Il pulsante li crea in Advertising, già etichettati con la fase, in pausa e pronti da rifinire. Per la fase 5 non si fanno annunci." },
      { h: "Il collegamento che vale di più", p: "Quando lo show up alle demo scende sotto il 60% e la spesa sulle fasi 2 e 3 è quasi a zero, nel Ciclo di vendita compare un avviso critico: le mancate presenze non nascono dal commerciale ma dal messaggio. Prima di aggiungere promemoria e telefonate di conferma, metti in campo gli annunci intermedi." },
    ],
  },
  {
    id: "metodo", t: "Il metodo in sette passi",
    intro: "Il framework a soglie dice quando spegnere un annuncio. Il metodo dice cosa costruire prima, ed è lì che si decide se la campagna funzionerà. Cinque passi prima di pubblicare, due dopo.",
    blocchi: [
      { h: "Il messaggio completo", p: "Il processo d'acquisto ha cinque fasi: riconoscimento del problema, raccolta di informazioni, valutazione delle soluzioni, acquisto, dopo l'acquisto. Portare le persone dalla fase 1 alla 4 con un solo annuncio è il messaggio corto, e logora le campagne in fretta. Assegna la fase a ogni ad set e annuncio: in Advertising trovi la spesa divisa per fase e un allarme se stai finanziando solo problema e acquisto." },
      { h: "Dati di cui fidarsi", p: "Nessun sistema è accurato al 100% e sistemi diversi non si riconcilieranno mai del tutto. Quello che serve è conoscere lo scarto: inserisci nel Ciclo di vendita le conversioni dichiarate dalle piattaforme accanto ai lead reali, e l'app calcola la precisione. Fra l'85 e il 95% si può comprare media con fiducia; uno scarto che cambia all'improvviso di solito è tracciamento rotto, non mercato cambiato." },
      { h: "Prima di pubblicare", p: "Piano di messaggio, piano di conversione, asset, tracciamento, test di validazione. La lista di installazione con i sette passi è nella Scheda di ogni cliente, con le caselle da spuntare: serve a vedere a colpo d'occhio quale cliente è partito senza fondamenta." },
      { h: "Dopo la pubblicazione", p: "Operare e scalare con regole chiare — le sei soglie — e ottimizzare lavorando solo sul punto di fallimento del percorso di conversione, che la sezione Funnel individua per te. Ritoccare tutto insieme è il modo migliore per non capire cosa ha funzionato." },
      { h: "I documenti", p: "In File trovi la sintesi di entrambi i documenti, framework delle decisioni e metodo in sette passi, con la fonte indicata. Se non ci sono, li inserisci dal riquadro in cima alla sezione o da Impostazioni → Contenuti pronti." },
    ],
  },
  {
    id: "semaforo", t: "Il sistema a semaforo",
    intro: "Due soli indicatori — quanto il costo reale si discosta dal KPI, e se il budget viene consumato — e tre risposte possibili: spegni, non fare nulla, scala. È il metodo AOA per decidere senza sensazioni.",
    blocchi: [
      { h: "I tre segnali", p: "Rosso: nessuna conversione, oppure costo oltre il 150% del KPI. Si spegne. Giallo: costo fra il 101% e il 150%, oppure costo nel KPI ma budget non consumato del tutto. Non si fa nulla — ed è un'istruzione, non un'assenza di istruzioni: non spegnere e non aumentare il budget. Verde: costo pari o sotto il KPI con budget consumato. Si scala." },
      { h: "Il KPI", p: "È il costo obiettivo per risultato, preso dai benchmark della campagna. Se la campagna raccoglie lead — obiettivo Lead, o lead già arrivati — il KPI è il CPL massimo e il costo reale si calcola sui lead; altrimenti è il CPA massimo e si calcola sulle conversioni. Senza quel numero il semaforo resta grigio, perché non c'è niente con cui confrontare il costo reale. Il semaforo scrive nel giudizio quale dei due sta usando." },
      { h: "Le fasi e i loro checkpoint", p: "Le fasi sono quelle che già usi, con i nomi del metodo: Test è il Proof of concept, Matura è il Validation period, Attiva è Scaling e mantenimento. Prima del checkpoint della propria fase una campagna resta grigia e non si giudica: in Proof of concept serve una spesa pari almeno al KPI — solida al doppio; in Validation servono i sette giorni di attesa; in Scaling si guarda ogni giorno." },
      { h: "Su quali numeri giudica", p: "Una campagna in fase Attiva viene valutata sulla finestra degli ultimi sette giorni, non sulla media dall'inizio: nell'intestazione di ogni scheda leggi \"valutata ultimi N giorni\" oppure \"dall'inizio\". È la differenza fra sapere come va adesso e sapere come è andata in media." },
      { h: "Il ritmo di spesa", p: "Il secondo indicatore è se la piattaforma consuma il budget giornaliero. Se non lo consuma, niente aumenti di budget: su una campagna che già non eroga non funzionerebbero. In quel caso si cresce in orizzontale, con nuovi pubblici o nuove campagne. Per l'aumento sulla stessa campagna — lo scaling verticale — il metodo chiede il budget pieno speso negli ultimi tre giorni consecutivi." },
      { h: "Come si usa la sezione", p: "A livello di agenzia scegli l'account dal menu; dentro un cliente è già quello. Sotto, l'elenco delle sue campagne con una casella ciascuna: ne analizzi una, alcune o tutte, e i pulsanti Tutte e Solo le attive fanno la scelta in un colpo. In cima al riquadro il conteggio per colore. Per ogni campagna scelta compare una scheda con costo, rapporto con il KPI, ritmo di spesa, speso, e la spiegazione del perché quel colore e cosa fare." },
      { h: "L'analisi 30 / 7 / 3", p: "Sotto il giudizio di ogni campagna ci sono tre righe: ultimi 30 giorni, ultimi 7, ultimi 3. Stessa campagna, stessi due indicatori, tre orizzonti. Serve a separare com'è andata da come sta andando: il totale dall'inizio è sempre in ritardo sulla realtà, perché i giorni buoni di tre settimane fa continuano a tenerlo su mentre la campagna sta già peggiorando. Ogni riga ha il suo pallino, il costo per risultato, la percentuale sul KPI e l'erogazione rispetto al budget giornaliero di quei giorni." },
      { h: "Come sono costruite le tre finestre", p: "Finiscono tutte ieri, mai oggi: il giorno in corso è parziale e affiancarlo a giorni interi abbasserebbe la spesa e alzerebbe i costi senza motivo. Sono mobili, quindi domani si spostano da sole di un giorno. Se la campagna è partita da meno di trenta — o di sette — giorni, la finestra parte dalla data di partenza reale e sotto l'etichetta compare quanti giorni copre davvero. I numeri si ricostruiscono dalle finestre compilate in proporzione ai giorni che cadono dentro l'intervallo: una settimana che sporge dal bordo conta solo per la sua parte, e quando succede la riga lo dichiara. Più le finestre sono corte e aggiornate, più l'analisi è precisa." },
      { h: "Leggere la direzione", p: "Sotto le tre righe l'app scrive cosa dicono messe in fila. Se il costo sale da 30 a 7 a 3 è il declino lento del metodo: la campagna sta già scivolando mentre la media la copre, e si interviene adesso invece di aspettare che il totale lo ammetta. Se scende nello stesso ordine, l'ultimo intervento sta funzionando e il totale lo mostrerà con qualche giorno di ritardo. Una finestra resta grigia finché la sua spesa non arriva almeno a una volta il KPI: su tre giorni capita spesso, ed è giusto così, perché su una spesa piccola il costo per risultato è rumore." },
      { h: "Il dettaglio su ad set e annunci", p: "Sotto ogni campagna c'è la stessa valutazione applicata ai suoi ad set e annunci, con il KPI della campagna: pallino, speso, quota sulla campagna, risultati, costo, rapporto con il KPI e CTR link, con gli annunci rientrati sotto il loro ad set. Il ritmo di spesa resta un fatto di campagna, quindi al livello inferiore si giudica solo il rapporto fra costo e obiettivo." },
      { h: "Prima di spegnere, togli il peso morto", p: "Una campagna può essere rossa perché il suo costo medio è inquinato da un annuncio che spende senza convertire. L'app allora ricalcola: toglie gli elementi rossi e mostra quanto costerebbe il risultato contando solo quelli che convertono, con quanto ha bruciato il peso morto. Se quel costo rientra nel KPI o diventa giallo, la mossa giusta è spegnere i rossi e lasciare che il budget si concentri su chi converte, invece di spegnere tutto. Se anche senza di loro resta oltre il 150%, la campagna va spenta davvero." },
      { h: "Cosa aspettarsi dalla concentrazione", p: "Spostare budget su un annuncio non ne abbassa il costo per risultato: quello resta il suo. Quello che cambia è il costo medio della campagna, che smette di pagare chi non produce. Nel caso tipico — un annuncio al 111% del KPI e uno a zero conversioni — spegnendo il secondo la campagna passa da rossa a gialla, non a verde: si lascia correre e si guarda al controllo successivo, non si scala." },
      { h: "Quale spegnere, quale scalare", p: "Sotto la tabella l'app scrive su cosa intervenire. Se la campagna è rossa ma contiene elementi verdi, dice di non spegnerla tutta e indica quali trascinano giù il risultato. Se è rossa dappertutto, dice che non c'è niente da salvare. Se è verde, indica chi sta producendo il risultato — cioè dove mettere il budget in più — e avverte se dentro c'è ancora qualcosa sopra il KPI, perché scalare una campagna che contiene un annuncio caro significa pagare di più anche quello. Se è gialla, spegnere un elemento rosso al suo interno è l'unico intervento ammesso, e spesso basta a riportarla nel KPI." },
      { h: "Quando un elemento resta grigio", p: "Sotto una spesa pari al KPI un ad set o un annuncio non viene giudicato: è lo stesso principio del checkpoint di campagna, applicato in piccolo. Serve a non spegnere una creatività che ha avuto venti euro di prova e tre impression." },
      { h: "Dove si vede l'attesa", p: "Una campagna in attesa lo dichiara dove la incontri: in Advertising, sotto lo stato, compare in ambra “in valutazione dopo modifica · ancora due giorni”, e passandoci sopra leggi cosa hai cambiato e fino a quando. Lo stesso nella tabella del cruscotto e nell'elenco del semaforo. Nel menu dei filtri c'è anche la voce “In valutazione dopo modifica”, per isolarle tutte in una volta." },
      { h: "Dopo una modifica si aspetta", p: "Quando intervieni su una campagna — spegni un annuncio, cambi il budget, tocchi il pubblico — il controllo successivo si fa tre giorni dopo, non prima. Nella scheda del semaforo il pulsante “Ho fatto una modifica” registra l'intervento con una nota su cosa hai cambiato: da quel momento la campagna resta grigia, con la data in cui riprendere, e compare una nota anche fra gli avvisi." },
      { h: "Correggerla o toglierla", p: "Quando una modifica è registrata, nella scheda compaiono tre pulsanti. “Correggi” cambia il testo e aggiorna anche la nota nel Diario. “Elimina” toglie l'attesa e il semaforo riprende subito a giudicare, chiedendoti se tenere la nota nel Diario — che è la storia dell'intervento — o cancellare anche quella. “Nuova modifica” sostituisce la precedente e fa ripartire i tre giorni. Cancellando la nota direttamente dal Diario, l'attesa sulla campagna si annulla da sé: un riferimento a una nota che non esiste più la terrebbe sospesa senza motivo." },
      { h: "La modifica finisce nel Diario", p: "Lo stesso intervento viene scritto nel Diario del cliente, collegato a quella campagna, con il tipo dedotto da come l'hai descritto: “spento AD1” diventa una nota di creatività, “alzato il budget” una nota di budget. Nella scheda del semaforo compaiono gli ultimi quattro interventi registrati su quella campagna, cliccabili per aprirli. Così il semaforo tiene l'ultimo stato e il Diario la storia." },
      { h: "Partire dal Diario", p: "Funziona anche al contrario: scrivendo una nota nel Diario e collegandola a una campagna, la casella “È un intervento: avvia i tre giorni di attesa” fa partire la sospensione del giudizio da lì. Utile quando annoti le cose a fine giornata invece che nel momento in cui le fai." },
      { h: "Perché proprio tre giorni", p: "Dopo un intervento l'algoritmo ridistribuisce la consegna, e i numeri dei primi giorni descrivono quella transizione, non la campagna. Giudicare lì dentro porta a disfare una modifica che non ha ancora avuto modo di produrre effetti. La stessa regola spiega perché conviene cambiare una cosa alla volta: se ne tocchi due, a fine attesa non sai quale ha prodotto il risultato." },
      { h: "La campanella alla scadenza", p: "Passati i tre giorni, accanto al nome della campagna compare una campanella ambra: la vedi in Advertising, nel cruscotto e nell'elenco del semaforo, e passandoci sopra leggi da quando la valutazione è pronta e cosa avevi cambiato. Nella colonna della fase la dicitura passa da “in valutazione dopo modifica” a “valutazione pronta da due giorni”." },
      { h: "Spegnere la campanella", p: "Nella scheda del semaforo compare il pulsante “Ho rivisto”: premendolo la campanella si spegne e l'avviso sparisce. Finché resta accesa significa che la decisione dopo quell'intervento non è ancora stata presa — ed è una cosa diversa dall'averla presa e aver deciso di non fare nulla. Fra gli avvisi compare una nota alla scadenza, che dopo tre giorni di silenzio diventa un'attenzione." },
      { h: "Elementi senza dati", p: "Nel dettaglio di una campagna, gli ad set o gli annunci senza spesa né impression non vengono mostrati: in fondo leggi quanti sono e come si chiamano. Di solito sono residui di importazioni fatte con nomi leggermente diversi, e affiancarli a quelli che lavorano confonde la lettura. Li trovi nella tabella Ad set e annunci, dove puoi eliminarli." },
      { h: "Il pallino in Advertising", p: "Lo stesso segnale compare come pallino colorato all'inizio di ogni riga nella tabella delle campagne: passandoci sopra leggi il giudizio e la ragione. Serve a vedere lo stato di tutto il portafoglio senza entrare nella sezione." },
      { h: "L'ordine delle decisioni", p: "Nella pratica il giro è sempre lo stesso. Primo: la campagna è in attesa dopo una modifica? Allora non si tocca. Secondo: ha superato il checkpoint della sua fase? Se no, si aspetta. Terzo: che colore ha? Se è rossa, si guarda dentro prima di spegnerla, perché il peso morto può bastare. Se è gialla, si spegne al massimo un elemento rosso al suo interno. Se è verde e il budget è consumato, si scala su chi produce. Quarto: qualunque cosa tu abbia fatto, registra l'intervento e ricontrolla fra tre giorni." },
      { h: "Come convive con gli avvisi", p: "Gli avvisi rispondono alla domanda \"cosa non va e perché\", con diagnosi su CTR, CPM, affaticamento, tracciamento e margine. Il semaforo risponde a una domanda diversa e più stretta: \"cosa faccio adesso con questa campagna\". Il primo serve a capire, il secondo a decidere in quindici minuti. Non si sostituiscono." },
    ],
  },
  {
    id: "framework", t: "Il framework decisionale",
    intro: "Gli avvisi non seguono soglie inventate: seguono il framework a sei soglie sequenziali di Jason Hornung, con l'aggiunta della lettura del margine. La prima soglia è un cancello, le altre valgono solo dopo.",
    blocchi: [
      { h: "1. Prima di giudicare, la spesa", p: "Una campagna va valutata solo dopo aver speso due volte il valore medio dell'ordine, o due volte il CPL obiettivo se raccoglie lead. Sotto quella cifra l'app mostra una nota grigia \"troppo presto per giudicare\" e sospende tutti gli altri giudizi: sono rumore. Imposta il valore medio ordine nella scheda cliente, altrimenti il cancello resta aperto." },
      { h: "2. CTR sotto 0,8%", p: "È un problema di creatività, non di pubblico. L'avviso ti dice di spegnere la creatività, tenere lo stesso pubblico e riscrivere il gancio con un angolo diverso: cambiare targeting qui è l'errore più comune." },
      { h: "3. CPM oltre il 40% sopra la media", p: "L'app calcola da sola la media del tuo account per quel cliente e confronta ogni campagna con quella. Un CPM alto in rapporto al tuo account significa che l'asta sta punendo la creatività, non che il pubblico è sbagliato." },
      { h: "4. Il cancello ROAS, CPL o CPA", p: "Sotto obiettivo dopo la soglia diventa critico solo se non c'è miglioramento nell'ultimo intervallo rilevato. Se invece sta risalendo, l'avviso resta di attenzione e ti dice di aspettare un intervallo — ma fissando la data in cui deciderai." },
      { h: "5. Slope test", p: "Tre intervalli consecutivi di peggioramento fanno scattare un avviso critico anche se i numeri complessivi reggono ancora: è l'affaticamento lento, quello che l'algoritmo maschera con lo storico. Funziona solo se salvi le rilevazioni con regolarità." },
      { h: "6. L'eccezione sulla frequenza", p: "Se la frequenza supera la soglia ma i risultati sono ancora a obiettivo, l'app non ti dice di spegnere: te lo segnala come nota e ti invita a lasciar correre. Spegnere per abitudine a frequenza 3 o 4 è uno dei modi più comuni di buttare via una creatività che funziona." },
      { h: "Applicare le soglie", p: "In Advertising il pulsante \"Soglie del framework\" imposta CTR minimo 0,80%, frequenza massima 4,50 e CPM massimo al 140% della media dell'account sulle campagne che stai guardando. ROAS e CPA obiettivo restano una tua decisione: dipendono dal margine, non da una regola generale. Il documento completo è in File." },
    ],
  },
  {
    id: "vendite", t: "Ciclo di vendita e chiusure",
    intro: "Qui vive la parte che l'advertising da solo non vede: cosa succede dopo che il lead è arrivato.",
    blocchi: [
      { h: "Fissate contro svolte", p: "Discovery e demo hanno due numeri distinti apposta. La differenza è lo show up, ed è quasi sempre il buco più costoso e il meno guardato." },
      { h: "Chi comanda sui numeri", p: "Se per quel mese esistono schede chiusura, sono loro a determinare chiusure, valore contratto e incassato: nella tabella compare \"da schede chiusura\". I campi manuali restano solo per i mesi che non dettagli." },
      { h: "Chiusure senza schede", p: "La sezione Chiusure conta i clienti chiusi, i contratti e l'incassato mese per mese con la stessa regola del Riepilogo mensile: in un mese con schede chiusura valgono le schede, altrimenti i numeri del Ciclo di vendita. Se carichi il ciclo da un file KPI o dal registro dei venditori, Chiusure, Panoramica e Riepilogo si aggiornano da soli. Le schede restano utili per sapere da quale campagna, ad set e annuncio arriva ogni cliente. In Panoramica i riquadri Costo per cliente, Contratti firmati e Ritorno usano l'ultimo mese che ha numeri commerciali, indicato accanto al titolo quando non è il mese in corso." },
      { h: "Fuori dai clienti, solo i totali", p: "Con Tutti i clienti selezionato, Ciclo di vendita, Chiusure e Advertising mostrano una riga per cliente con i totali: mesi registrati, lead, demo, chiusure, contratti e incassato; spesa, CPL e ROAS per l'advertising. Ogni importo resta nella valuta del suo cliente. I dettagli, cioè i singoli mesi, le schede chiusura, le campagne, gli ad set e le finestre, stanno dentro il cliente: un clic sulla riga ti porta lì." },
      { h: "Importare da un file KPI", p: "Se il cliente tiene i suoi numeri in un foglio di riepilogo, dentro il cliente, nella scheda File o nel Ciclo di vendita, carichi il file Excel o CSV nel riquadro Importa i numeri da un file KPI. Il CRM cerca le tabelle con le voci in colonna (spesa, lead, appuntamenti fissati e svolti, nuovi clienti, fatturato e incassato dei nuovi clienti) e i mesi in riga, anche con una colonna per funnel sotto ogni mese, e te le elenca: scegli quali caricare e in quale cliente o sottocliente. Più tabelle si sommano mese per mese, utile quando un mese ha i funnel separati e nessuna colonna totale. Gli appuntamenti diventano demo fissate e svolte, i nuovi clienti chiusure. Prima di scrivere vedi l'anteprima mese per mese; si aggiornano solo i campi presenti nel file, e ricaricarlo sostituisce i numeri invece di sommarli." },
      { h: "Contratto e incassato", p: "Sono due cose diverse e vanno tenute separate: sul contratto si festeggia, sull'incassato si paga. La quota incassata e il residuo sono calcolati da soli." },
      { h: "Capacità", p: "Slot disponibili e ore al primo contatto servono a capire se il collo di bottiglia è il traffico o la capacità di lavorarlo. Sopra le dodici ore di attesa parte un avviso: prima di alzare il budget, riduci quel tempo." },
    ],
  },
  {
    id: "linee", t: "Linee di business",
    intro: "Quando un cliente fa pubblicità per due attività diverse — due offerte, due mercati, due margini — sommarle in un totale solo produce numeri che non descrivono nessuna delle due.",
    blocchi: [
      { h: "Assegnare la linea", p: "Nella scheda della campagna, sotto lo stato, c'è il campo \"Linea di business\": scrivi il nome che usi tu — per esempio Corsi e Consulenza, oppure Ristorante e Catering. Il campo suggerisce le linee già usate, così restano scritte allo stesso modo. Le campagne senza linea finiscono in un gruppo a parte." },
      { h: "Leggerle separate", p: "In Advertising compare la tabella \"Totali per linea di business\", con spesa, CTR link, lead, CPL, conversioni, CPA, ROAS e contratti di ciascuna. Non c'è la riga del totale, ed è voluto. Cliccando una riga isoli quella linea in tutta la sezione: tabella, riepiloghi, ritmo di erogazione." },
      { h: "L'avvertimento", p: "Se un cliente ha più di una linea e non ne hai isolata nessuna, sopra la tabella compare una nota che ricorda che i numeri qui sopra le stanno sommando. Serve soprattutto per la soglia di valutazione: il valore medio dell'ordine di una linea non vale per l'altra, e il framework si regge su quel numero." },
      { h: "Quando invece conviene un sottocliente", p: "La linea di business separa le campagne dentro lo stesso cliente. Se le due attività hanno anche commerciali, funnel e contratti distinti — e vuoi un ciclo di vendita e delle chiusure separate — allora la struttura giusta è il sottocliente, non la linea." },
    ],
  },
  {
    id: "mensile", t: "Riepilogo mensile",
    intro: "Un'unica tabella che mette insieme quello che spendi e quello che entra, mese per mese. C'è a livello di agenzia, dentro ogni cliente e dentro ogni sottocliente.",
    blocchi: [
      { h: "Cosa contiene", p: "Per ogni mese: spesa pubblicitaria, lead, costo per lead, demo svolte, chiusure, costo per chiusura, valore dei contratti firmati, incassato e ritorno. In fondo la riga con i totali del periodo, con costo per lead e per chiusura ricalcolati sui totali e non mediati tra i mesi." },
      { h: "Da dove arrivano i numeri", p: "La spesa viene dalle finestre delle campagne quando le hai compilate, altrimenti da quella dichiarata nel ciclo di vendita, e nei mesi in cui usa la seconda fonte te lo scrive sotto il mese. Lead e demo arrivano dalle giornate dei venditori se le hai caricate, altrimenti dal periodo del ciclo di vendita. Contratti e incassato dalle schede chiusura." },
      { h: "Il ritorno non è il ROAS", p: "La colonna Ritorno è il valore dei contratti firmati diviso la spesa pubblicitaria, non il ROAS dichiarato dalla piattaforma. È il numero che porti in riunione, perché è l'unico che il cliente riconosce come suo." },
      { h: "A livello di agenzia", p: "Sotto la tabella dei mesi trovi anche il quadro per cliente su tutto il periodo, ordinato per valore dei contratti, con i sottoclienti indicati. È la lettura che ti dice dove sta davvero il tuo lavoro, che non sempre coincide con dove sta la spesa." },
      { h: "Esportare", p: "Il pulsante Esporta in Excel genera un .xlsx con la tabella e la riga dei totali, pronto da allegare a un rapporto o da mandare al cliente." },
    ],
  },
  {
    id: "elenchi", t: "Importare un elenco di clienti",
    intro: "Quando un cliente ti porta i suoi clienti finali — gli avvocati di una società di consulenza, i professionisti di un gruppo — puoi crearli tutti in una volta da un foglio invece che uno per uno.",
    blocchi: [
      { h: "Preparare il foglio", p: "In Google Fogli o Excel serve una riga per cliente e una prima riga di intestazioni. L'unica colonna indispensabile è il nome. Le altre sono facoltative e vengono riconosciute da sole: settore, sede, stato, compenso mensile, budget adv, valore medio ordine, margine, referente, email, ragione sociale, partita IVA e note." },
      { h: "Portarlo nell'app", p: "Da Google Fogli: File → Scarica → Valori separati da virgola, oppure Microsoft Excel. Il collegamento diretto al foglio non funziona, perché Google non consente la lettura da un'altra pagina: serve il file scaricato. Poi Importa dati, tipo \"Clienti e sottoclienti\", e carica." },
      { h: "Dove finiscono", p: "Nel secondo menu in alto scegli il cliente madre: ogni riga diventa un suo sottocliente, con scheda, campagne, funnel, messaggio e ciclo di vendita propri. Lasciando “come clienti principali” vengono creati al primo livello. Chi ha già lo stesso nome sotto lo stesso madre viene aggiornato invece che duplicato, quindi puoi ricaricare il foglio aggiornato senza paura." },
      { h: "La scorciatoia", p: "Dal cruscotto di un cliente, nella scheda Sottoclienti, il pulsante \"Importa elenco\" apre direttamente l'importazione già impostata su quel cliente madre." },
    ],
  },
  {
    id: "registro", t: "Registro dei venditori",
    intro: "Un modulo che i venditori compilano una riga al giorno, e che tu carichi una volta a settimana senza reinserire nulla.",
    blocchi: [
      { h: "Il modulo", p: "Dalla sezione Ciclo di vendita, dentro il cliente, il pulsante \"Scarica il modulo Excel\" genera un file .xlsx con due fogli, Registro e Istruzioni, e le colonne già pronte: data, venditore, funnel di provenienza, lead ricevuti, chiamate, discovery fissate e svolte, demo fissate e svolte, offerte, chiusure, valore contratto, incassato, ore al primo contatto e note. Si apre in Excel o Fogli Google e si compila una riga per giorno per venditore." },
      { h: "Il caricamento", p: "Dallo stesso riquadro carichi il file compilato, in CSV o Excel. L'app riconosce le colonne dai nomi del modulo, ti dice quante giornate e quanti venditori ha trovato e su quali mesi, e scrive solo dopo la conferma. Se un mese non esisteva ancora nel ciclo di vendita, viene creato." },
      { h: "Se carichi un file diverso", p: "Se nel riquadro del registro carichi un foglio di riepilogo, con i mesi in riga invece dei giorni, il CRM lo riconosce e lo passa al riquadro Importa i numeri da un file KPI, subito sopra." },
      { h: "Niente doppioni", p: "Le righe sono identificate da data, venditore e funnel: ricaricare lo stesso file o una versione aggiornata sostituisce le giornate corrispondenti invece di sommarle. Puoi quindi caricare ogni settimana il file completo dall'inizio del mese senza pensarci." },
      { h: "Cosa cambia nei numeri", p: "Quando un mese ha giornate caricate, lead, chiamate, discovery, demo e offerte di quel periodo arrivano da lì e non dai totali scritti a mano, e le ore al primo contatto diventano la media delle giornate compilate. Chiusure e importi continuano a venire dalle schede della sezione Chiusure quando ci sono." },
      { h: "Le due letture in più", p: "Con le giornate caricate compaiono due tabelle: per venditore, con show up e tasso di chiusura di ciascuno, e per funnel di provenienza dei lead chiamati. La seconda va confrontata con la provenienza registrata nelle schede chiusura: quando i due quadri non coincidono, di solito è il tracciamento a mentire, non il venditore." },
    ],
  },
  {
    id: "analisi", t: "Grafici, pianificatore, funnel",
    intro: "Tre strumenti che usano i dati già inseriti per farti decidere invece che constatare.",
    blocchi: [
      { h: "Creare un modello da zero", p: "Nella scheda Modelli di percorso il pulsante \"+ Nuovo modello\" apre un modulo con nome e sequenza: componi i passi e lo salvi, senza dover prima costruire un funnel vero. Cliccando su un modello esistente lo riapri e lo modifichi. I modelli non hanno numeri: sono solo la struttura." },
      { h: "Le campagne nel modulo", p: "Creando o modificando un funnel, il campo \"Campagne che portano traffico\" elenca le campagne del cliente con una casella ciascuna, più le scorciatoie Tutte, Solo le attive e Nessuna. Accanto a quelle già collegate altrove leggi su quale funnel si trovano, perché spuntandole le sposti." },
      { h: "Modelli tuoi", p: "Oltre ai quattro percorsi predefiniti puoi salvare i tuoi: quando un funnel ha la sequenza che ti serve, il pulsante \"Salva questo come modello\" nella scheda Modelli di percorso la conserva — nomi e tipo dei passi, senza i numeri — e la ritrovi nel modulo del funnel sotto i predefiniti, pronta da applicare a un altro cliente. Salvando con un nome già usato lo sostituisci. Eliminare un modello non tocca i funnel già creati con quello." },
      { h: "Riempire i passi dai dati", p: "Se il funnel ha campagne collegate, sopra la tabella compare il pulsante “Riempi dai dati del CRM”, con accanto cosa sta per scrivere. Mette le impression e i clic sul link nel passo Annuncio, i clic come ingressi della landing, i lead come moduli inviati, e — se hai compilato il ciclo di vendita o caricato il registro venditori — appuntamenti fissati e svolti nei passi Prenotazione e Chiamata. Non tocca i passaggi che vivono solo sulla pagina, come quante persone dalla landing arrivano al modulo: quelli restano tuoi, dalle statistiche del sito." },
      { h: "Come leggere le colonne", p: "Entrano è quante persone arrivano a quel gradino, Passano quante proseguono al successivo. Tasso e Persi si calcolano da sé. La regola che tiene insieme il funnel è che chi passa da un gradino entra nel successivo: se i due numeri non coincidono l'app lo segnala sotto il nome, perché significa che stai misurando due cose diverse." },
      { h: "Collegare le campagne", p: "Nella scheda del funnel c'è l'elenco delle campagne del cliente con una casella ciascuna: spunti quelle che portano traffico a quel percorso. In alto due scorciatoie, \"Collega tutte\" e \"Solo le attive\", e l'intestazione dice quante sono collegate e quanta spesa rappresentano. Una campagna appartiene a un funnel solo: se era collegata altrove, spuntandola qui viene spostata, e accanto al nome leggi su quale funnel si trovava." },
      { h: "Grafici", p: "Scegli ambito, raggruppamento, metrica e se vederla a barre o a torta. Quando la metrica è un rapporto — ROAS, CTR, tasso di chiusura — la torta viene sostituita dalle barre: le fette rappresentano quote di un totale, e i rapporti non si sommano." },
      { h: "Pianificatore", p: "Metti l'obiettivo di chiusure e risale con i tuoi tassi reali fino a lead e budget necessari. È lo strumento con cui si giustifica un aumento di budget davanti a un cliente." },
      { h: "Funnel", p: "Descrivi la sequenza passo per passo con i numeri di ciascuno. L'app trova il passaggio con la resa più bassa e ti dà una diagnosi diversa a seconda che il buco sia sulla landing, sul video, sul modulo o sul calendario. Da lì puoi generare una pagina HTML da caricare dove vuoi." },
    ],
  },
  {
    id: "calendario", t: "Calendario",
    intro: "Un piano unico dove finiscono le scadenze dei task, gli avvii e le chiusure delle campagne, le finestre di analisi, i clienti chiusi e gli eventi che aggiungi tu.",
    blocchi: [
      { h: "Quattro viste", p: "Giorno con le ore, settimana, mese e anno con i dodici calendarietti, dal menu in alto a destra. Le frecce spostano di un giorno, una settimana, un mese o un anno a seconda di dove sei, e “Oggi” riporta alla data corrente. Nella vista anno ogni mese mostra quante voci contiene e i colori di cui è fatto: cliccandolo ci entri." },
      { h: "La giornata, ora per ora", p: "La vista giorno ha in testa la fascia “Tutto il giorno”, dove finiscono scadenze, avvii di campagna e chiusure — cose che non hanno un orario — e sotto la griglia delle ore. Gli eventi con un orario si posizionano nella loro fascia, e il campo Durata in ore ne determina l'altezza. Sulla data di oggi una linea bordeaux segna l'ora corrente. Il doppio clic su una fascia, o il + che compare passandoci sopra, crea un evento già all'ora giusta." },
      { h: "L'anno, dodici mesi veri", p: "La vista anno mostra i dodici mesi come calendarietti completi, non come riquadri riassuntivi: ogni giorno che contiene qualcosa ha sotto fino a tre puntini con i colori di ciò che c'è, e passandoci sopra leggi cosa. Cliccando un giorno entri nella sua vista giornaliera, cliccando il nome del mese nella vista mensile. Accanto al nome, quante voci contiene quel mese." },
      { h: "Cosa compare, e come lo scegli", p: "La legenda sotto l'intestazione è fatta di caselle: accendi e spegni eventi, task, campagne, finestre di analisi e clienti chiusi. Le finestre partono spente perché su clienti con molte campagne riempiono il calendario; accendile quando pianifichi le rilevazioni." },
      { h: "I colori", p: "Gli eventi che crei tu prendono il colore della loro categoria, e le categorie le definisci nella scheda in fondo: nome e colore, con il significato che decidi tu — lancio campagna, revisione dati, call, scadenza, obiettivo. Le altre voci hanno colori propri: i task seguono la priorità, dal bordeaux dell'urgente all'ambra della media, e diventano verdi quando sono fatti." },
      { h: "Aggiungere", p: "Il pulsante “+ Evento” crea sulla data che stai guardando; dentro la griglia, il + che compare passando sul giorno — o un doppio clic sulla cella — crea direttamente lì. Un evento può durare più giorni: compila anche il campo “Al” e occuperà tutte le date intermedie. Cliccando una voce esistente la apri e la modifichi, che sia un evento o un task." },
      { h: "Il legame con gli avvisi", p: "Un task scaduto e ancora aperto diventa un avviso: critico se la priorità è urgente o alta, di attenzione negli altri casi. Un task urgente in scadenza oggi o domani compare come nota. Così le scadenze non vivono solo in una lista che puoi non aprire: entrano nello stesso posto dove guardi i problemi delle campagne." },
      { h: "Dove si trova", p: "C'è a livello di agenzia, dove vedi tutti i clienti con il nome accanto a ogni voce, e dentro ogni cliente, dove resta solo il suo. Gli eventi possono essere legati a un cliente o riguardare tutta l'agenzia." },
    ],
  },
  {
    id: "ada", t: "Le regole del framework",
    intro: "In cima alla sezione Automazioni c'è il pacchetto del framework decisionale: undici regole, una per ogni soglia, che trasformano gli avvisi in task appena si accendono.",
    blocchi: [
      { h: "Cosa contiene", p: "Gancio debole sul CTR link, asta cara sul CPM, spesa senza conversioni, sotto il pareggio sul margine, ritorno sotto obiettivo su ROAS, CPL e CPA, peggioramento costante sul test di tendenza, pubblico saturo sulla frequenza, erogazione anomala, e una rete di sicurezza che prende qualunque avviso critico anche su metriche non coperte dalle altre." },
      { h: "Come si installa", p: "Scegli dal menu se applicarle a tutti i clienti o a uno solo, poi premi Installa. La tabella mostra per ciascuna soglia quando scatta, che task crea e dove è già installata. Le regole già presenti non vengono duplicate, quindi puoi aggiungere un cliente alla volta senza fare danni." },
      { h: "Le soglie non sono scritte qui", p: "Ogni regola si appoggia agli avvisi, che usano i benchmark della singola campagna. Impostando un CTR minimo diverso su un cliente, la stessa regola scatta a un valore diverso. È il motivo per cui il pacchetto funziona su clienti con economie opposte senza doverlo duplicare — ma anche il motivo per cui su una campagna senza benchmark impostati resta muto." },
      { h: "Le priorità", p: "Non sono tutte uguali, e seguono il danno che fa il problema: urgente per spesa senza conversioni, perdita sotto il pareggio e avvisi critici; alta per gancio, ritorno, costo e tendenza; media per asta e frequenza, che raramente richiedono di agire lo stesso giorno." },
      { h: "Il limite", p: "Come tutte le automazioni, queste regole girano solo mentre l'app è aperta. Non vigilano al posto tuo di notte: ti risparmiano la fatica di cercare i problemi, non quella di aprire il software." },
    ],
  },
  {
    id: "lavoro", t: "Task e automazioni",
    intro: "La parte che tiene in ordine l'esecuzione.",
    blocchi: [
      { h: "Task", p: "Bacheca per stato o elenco per scadenza, con aree di lavoro, priorità, stime e passaggi. I ricorrenti si rigenerano da soli quando li chiudi: comodo per report e controlli mensili." },
      { h: "Automazioni", p: "Una regola collega un innesco a un'azione. Gli inneschi guardano avvisi, soglie, budget, ritardi, chiusure e mesi senza dati; le azioni creano un task o chiamano un webhook. Ogni evento scatta una volta sola e finisce nel registro." },
      { h: "Limite da conoscere", p: "Le regole girano mentre l'app è aperta. Non esiste modo di farle partire a browser chiuso, e non si possono ricevere dati dall'esterno: servirebbe un server acceso con un indirizzo pubblico." },
    ],
  },
  {
    id: "aggiornare", t: "Aggiornare i dati da Meta",
    intro: "Ogni mattina, all'ora scelta in Impostazioni (predefinita le 9, nel fuso orario dell'agenzia), il CRM scarica da Meta Ads i dati fino a ieri per tutti gli account collegati ai clienti: campagne, ad set, annunci e finestre di valutazione. Non serve importare file.",
    blocchi: [
      { h: "Una volta sola: collegare Meta", p: "In Impostazioni, nel riquadro Piattaforme pubblicitarie, premi Collega a Meta e accedi con il tuo profilo Facebook: il CRM ottiene un accesso in sola lettura agli account pubblicitari che gestisci, senza bisogno di un Business Manager unico. Meta fa durare l'accesso 60 giorni: una settimana prima della scadenza il riquadro diventa rosso e basta premere Rinnova il collegamento." },
      { h: "Fuso orario e ora dell'aggiornamento", p: "Nel riquadro Fuso orario delle Impostazioni scegli il fuso dell'agenzia e l'ora a cui il CRM scarica i dati di Meta ogni giorno (predefinita le 9). Per uno spazio di lavoro nuovo il CRM propone il fuso del computer di chi lo crea: un'agenzia di New York riceve i dati alle 9 di New York, non alle 3 di notte. Lo stesso fuso vale per il riepilogo e i promemoria del bot Telegram. I dati di ogni account restano divisi per giorno secondo il fuso dell'account pubblicitario, come in Gestione inserzioni." },
      { h: "Poi: collegare gli account ai clienti", p: "Nello stesso riquadro, Aggiorna l'elenco degli account mostra tutti gli account a cui il tuo profilo ha accesso. Accanto a ogni cliente il menu parte da «scegli un account»: aprendolo, in cima trovi i Suggeriti per nome, quelli che gli somigliano (anche per iniziali: BDF per Boutique Del Fitness), e sotto tutti gli altri. Scegli e premi Collega. Lo stesso si fa dalla Scheda cliente. Un cliente può avere più account; un account appartiene a un solo cliente." },
      { h: "Cosa arriva", p: "Le campagne attive o che hanno speso negli ultimi 90 giorni, con stato, budget giornaliero, date, spesa, impression, clic, clic sul link, lead, acquisti e ricavi. Sotto ogni campagna i suoi ad set e annunci con i totali. Le finestre Test, Matura e le settimane Attive sono già sommate da Meta, con la frequenza vera di ciascuna." },
      { h: "Come si agganciano", p: "Una campagna che avevi già scritto a mano si aggancia per nome la prima volta; da lì in poi per l'identificativo di Meta, quindi se la rinomini su Meta resta la stessa. Le campagne nuove nascono da sole sotto il cliente, con l'obiettivo dedotto da Meta. Le finestre di Meta prendono il posto di quelle scritte a mano che si sovrappongono; note e importi di contratto restano." },
      { h: "Cosa resta tuo", p: "Lo stato In test, le soglie, la linea, il funnel, le note, lo stadio dei creativi e, se l'hai segnato come scritto a mano, l'esito commerciale (conversioni e ricavi). Il budget giornaliero cambiato su Meta entra nello storico con la data del giorno, come quando lo cambi dalla scheda." },
      { h: "Aggiorna ora", p: "Se non vuoi aspettare il mattino, Aggiorna ora scarica subito tutti i clienti collegati. Per uno solo usa il pulsante Aggiorna sulla sua riga in Impostazioni, oppure dalla sua scheda: se ha sottoclienti trovi due pulsanti, Aggiorna solo questo cliente (i suoi account e basta) e Aggiorna cliente e sottoclienti: utile quando un account si è fermato o ti serve un dato fresco. Con molti account ci vuole qualche minuto: puoi continuare a lavorare, i numeri arrivano alla fine. Meta ritocca gli ultimi giorni per qualche tempo, per questo ogni aggiornamento riscarica anche la settimana precedente." },
      { h: "Scegliere quali campagne scaricare", p: "Nella Scheda cliente, sotto ogni account collegato, c'è Campagne da scaricare. Normalmente il CRM prende le attive e quelle che hanno speso negli ultimi 90 giorni. Con Solo le attive prende solo quelle accese (quando una si spegne, la registra un'ultima volta in pausa e poi la lascia). Con Solo quelle che scelgo il CRM legge da Meta l'elenco delle campagne dell'account, accese prima e poi per spesa, e scarichi solo quelle che spunti; Cambia la scelta riapre l'elenco. Le campagne non scelte già presenti nel CRM restano dove sono: se non ti servono le elimini e non verranno ricaricate. La scelta vale dal prossimo aggiornamento; per vederla subito premi Aggiorna questo cliente." },
      { h: "Account che si aggiornano solo a mano", p: "Sempre sotto l'account, togliendo la spunta Nell'aggiornamento automatico l'account resta fuori dall'aggiornamento delle 9 e da Aggiorna ora: si aggiorna solo quando premi Aggiorna sulla sua riga in Impostazioni o Aggiorna questo cliente nella sua scheda. Serve per gli account pesanti, quelli fermi o quelli che ti basta caricare una volta: scarichi i dati, togli la spunta, e l'aggiornamento di tutti diventa più veloce. In Impostazioni, accanto al nome dell'account, compare la sua regola (solo attive, campagne scelte, solo a mano)." },
      { h: "Quando qualcosa non va", p: "Un account con un errore (pagamento in sospeso, accesso tolto, token scaduto) compare in rosso nel riquadro delle Impostazioni con il messaggio di Meta; gli altri si aggiornano lo stesso. Le campagne accese su un account con pagamento in sospeso passano in Errore pagamento; su un account disattivato, in revisione o chiuso passano rispettivamente in Account disattivato, Account in revisione o Account chiuso, con l'avviso critico e i giorni di fermo come per il pagamento. Sotto lo stato di ogni campagna dell'account compare in rosso anche lo stato su Meta, con il motivo della disattivazione quando Meta lo dà (per esempio violazione delle norme pubblicitarie o account compromesso). Le campagne che erano accese diventano ferme, con l'avviso; quelle che avevi già spento restano In pausa o Conclusa, con la scritta campagna già spenta: l'account non le ha fermate, ma se le riaccendi non erogheranno." },
      { h: "Il pulsante Chiedi i dati aggiornati", p: "Resta per i clienti che non hai collegato o per le piattaforme diverse da Meta: prepara la richiesta da incollare in chat, come prima." },
    ],
  },
  {
    id: "telegram", t: "Il bot Telegram",
    intro: "Un assistente su Telegram che crea task e note nel diario quando glielo scrivi, risponde alle domande sui dati del CRM e ogni mattina ti manda il riepilogo dei task.",
    blocchi: [
      { h: "Collegarlo", p: "In Impostazioni, riquadro Bot Telegram, premi Collega Telegram e apri il link che compare: Telegram apre il bot, premi Avvia e la chat è collegata. Il link vale 30 minuti e funziona una volta sola. Il bot risponde solo alle chat collegate così; con chiunque altro dice che è privato." },
      { h: "Cosa puoi chiedergli", p: "Scrivi come scriveresti a un collaboratore: «crea un task per BDF per venerdì: report di ottobre, priorità alta», «segna nel diario che oggi abbiamo spento le campagne Rottamazione», «sposta la call con Oliva a lunedì alle 15», «che scadenze ho domani?», «quali campagne sono ferme?». Capisce le date a parole, i nomi abbreviati dei clienti e più richieste nello stesso messaggio; se il cliente è ambiguo chiede quale. Se sbaglia, correggilo: «no, la data è il 24»." },
      { h: "Quando compare nel CRM", p: "Il bot non scrive direttamente nei dati: mette le richieste in coda e il CRM aperto le inserisce entro un minuto, oppure appena lo apri. Così nulla di quello che stai facendo nel CRM viene sovrascritto. La conferma su Telegram invece arriva subito." },
      { h: "Il riepilogo del mattino", p: "All'ora che scegli, nel fuso orario dell'agenzia (predefinita 9:10, dopo l'aggiornamento di Meta) arrivano i task in ritardo, quelli di oggi, quelli di domani a priorità alta o urgente e le campagne ferme. Sotto, per ogni task di oggi o in ritardo, i pulsanti Fatto e Domani. Nel weekend solo se lo attivi." },
      { h: "Task con orario", p: "Se un task ha un orario (campo Ora nella scheda del task, o detto al bot), il bot ti avvisa prima, con il preavviso che scegli, e ti dà i pulsanti Fatto, Domani e Lunedì. Chiudere da Telegram un task che si ripete fa nascere il successivo, come la spunta nel CRM." },
      { h: "Cosa serve", p: "Due segreti nelle Edge Functions di Supabase: TELEGRAM_TOKEN, il token del bot creato con @BotFather, e ANTHROPIC_API_KEY, la chiave con cui il bot usa Claude per capire i messaggi. I vocali arriveranno in un secondo momento." },
    ],
  },
  {
    id: "academy", t: "Academy",
    intro: "La sezione Academy, con il cappello di laurea nel menu, porta nel CRM i corsi dell'Accademia dell'Advertising: lezioni gratuite da guardare subito, corsi a pagamento da sbloccare.",
    blocchi: [
      { h: "Il catalogo", p: "Ogni corso ha la sua scheda con livello, descrizione, numero di lezioni e avanzamento. «Gratis» vuol dire che è libero; il lucchetto con il prezzo che va sbloccato; alcuni corsi a pagamento hanno una lezione di anteprima gratuita. I filtri mostrano tutti i corsi, solo quelli gratuiti o solo i tuoi. Finché i corsi veri non sono collegati vedi un catalogo di esempio, segnalato in cima." },
      { h: "Guardare e sbloccare", p: "Una lezione libera si apre con il suo video e la casella per segnarla come seguita. Una lezione bloccata mostra il pulsante per sbloccare il corso, che porta alla pagina di acquisto dell'Accademia indicata in fondo alla sezione." },
      { h: "Suggerimenti nel CRM", p: "In Advertising e in Winner ads compare la lezione dell'Academy più utile per quella sezione, con il pulsante per guardarla o scoprirla." },
    ],
  },
  {
    id: "winner", t: "Winner ads e connettore per Claude",
    intro: "Il Claude di un cliente (per Boutique del Fitness quello collegato a Hyros, con tutto il second brain) salva nel CRM i numeri di ogni annuncio e le sue analisi. Il CRM ne fa la classifica sui tuoi benchmark, nella sezione Winner ads del cliente.",
    blocchi: [
      { h: "Creare il connettore", p: "In Impostazioni, riquadro Connettore per Claude, premi Crea il connettore: compaiono l'indirizzo del server e la chiave, che si vede una volta sola. Su claude.ai apri Impostazioni → Connettori → Aggiungi connettore personalizzato, chiamalo «CRM NC Media», incolla l'indirizzo, lascia «Nessun accesso» e in Intestazioni della richiesta aggiungi Authorization con valore «Bearer» seguito dalla chiave: così la chiave resta cifrata in claude.ai e non sta nell'indirizzo. Nuovo indirizzo sostituisce la chiave (quello vecchio smette subito di funzionare e va cambiato anche su claude.ai); Metti in pausa blocca il connettore senza eliminarlo." },
      { h: "Il Progetto del cliente", p: "Copia le istruzioni per il Progetto e incollale nelle istruzioni del Progetto di Boutique del Fitness su claude.ai. Da lì basta chiedere «aggiorna le winner ads nel CRM»: Claude legge i benchmark del CRM, prende da Hyros i numeri di ogni annuncio, li salva, scrive la sua analisi e, se serve, crea i task. Per farlo ogni mattina senza scrivere niente, crea un'attività in Pianificate su claude.ai dentro quel Progetto, con la stessa richiesta." },
      { h: "La classifica", p: "Per ogni annuncio il CRM calcola il costo per lead, per appuntamento e per cliente, li divide per i benchmark (predefiniti 9 €, 20 € e 100 €) e fa la media: è il punteggio. Sotto 1 l'annuncio batte i riferimenti, più è basso meglio è. In verde i costi sotto il benchmark, in rosso quelli oltre il 30% sopra. I filtri, cioè minimo di clienti (predefinito 5, perché un annuncio con due clienti può essere primo per caso), minimo di spesa, pubblico e solo attive, valgono per il cliente e li legge anche Claude." },
      { h: "Ordinare secondo l'obiettivo", p: "Il menu Ordina per, o un clic sull'intestazione di una colonna, cambia il criterio della classifica: equilibrio dei tre costi (il punteggio, predefinito), costo per lead, per appuntamento o per cliente più basso, più clienti, più appuntamenti, più lead, chiusura più alta, spesa o incassato più alti. Se cerchi lead a basso costo ordini per CPL, se vuoi scalare chi porta clienti ordini per clienti o costo per cliente. Gli annunci senza il dato (per esempio nessun appuntamento quando ordini per costo per appuntamento) vanno in fondo. La scelta resta salvata per il cliente; i filtri continuano a valere." },
      { h: "Freddo e retargeting, copie raggruppate", p: "Le classifiche sono due, una per il traffico freddo e una per il retargeting, perché i costi dei due pubblici non si confrontano. Con «raggruppa le copie» (acceso di base) le copie della stessa creatività in campagne o ad set diversi, anche in pausa, diventano una riga sola con i numeri sommati: così una creatività appena duplicata non sparisce perché la copia nuova ha uno o due giorni di vita. La creatività è riconosciuta dal codice che Claude indica, altrimenti dal numero all'inizio del nome (71 per «71 - tempo perso»), altrimenti dal nome senza «copia». Con il raggruppamento, «solo con una copia accesa» tiene le creatività che hanno almeno una copia attiva; sotto il nome leggi quante copie ci sono e quante sono accese." },
      { h: "Fotografie e andamento", p: "Ogni salvataggio è la fotografia di un giorno: dal menu sopra la classifica scegli quale vedere. Accanto alla posizione la freccia dice quanti posti ha guadagnato o perso rispetto alla fotografia precedente; «nuovo» vuol dire che prima non era in classifica. Salvare di nuovo lo stesso giorno sostituisce la fotografia." },
      { h: "Le analisi", p: "Sotto la classifica trovi le analisi scritte dal Claude del cliente, la più recente aperta. Restano anche quando cambiano i numeri, così puoi rileggere cosa diceva un mese fa. Dentro un cliente con sottoclienti, in consolidato vedi anche quelle dei sottoclienti." },
      { h: "Cosa può fare Claude con il connettore", p: "Leggere l'elenco dei clienti, i dati di un cliente (campagne, ciclo di vendita, task aperti) e le analisi precedenti; salvare numeri degli annunci e analisi; creare task. Non può modificare né cancellare il resto del CRM." },
    ],
  },
  {
    id: "integrazioni", t: "Task da altri software",
    intro: "Un software esterno, per esempio Boschetto di Kesia, manda al CRM le task che ti assegna: le trovi nei Task sotto il sottocliente giusto e, quando le cambi qui, il cambio torna al software di origine.",
    blocchi: [
      { h: "Collegarlo", p: "In Impostazioni, riquadro Software esterni: scrivi il nome del software, scegli il cliente principale (per Kesia, il suo cliente) e premi Collega un software. Compare una chiave che si vede una volta sola: copia le istruzioni con la chiave e mandale a chi gestisce il software. Se la chiave va persa o finisce nelle mani sbagliate, Nuova chiave la sostituisce e la vecchia smette subito di funzionare." },
      { h: "Il sottocliente giusto", p: "Ogni task porta con sé il nome del cliente o progetto che ha nel software di origine. Quei nomi compaiono nel riquadro, sotto «Clienti di …», man mano che arrivano le prime task: accanto a ciascuno scegli il sottocliente del CRM. Le task già arrivate con quel nome si spostano subito; quelle senza abbinamento restano sul cliente principale. Se sposti a mano una singola task su un altro sottocliente, gli aggiornamenti successivi non la riportano indietro." },
      { h: "Andata", p: "Quando nel software una task viene creata o modificata, arriva al CRM e il CRM aperto la inserisce entro un minuto, come i comandi del bot Telegram. Rimandare la stessa task la aggiorna, non la duplica. Una task eliminata alla fonte viene segnata come fatta. Nei Task ha l'etichetta «da Boschetto»: se il software manda il link, cliccandola apri la task nell'originale." },
      { h: "Ritorno", p: "Quando cambi una di queste task nel CRM (stato, scadenza, ora, priorità o titolo), il CRM lo comunica dopo pochi secondi. Il software può leggere i cambi quando vuole, oppure darti un suo indirizzo https: lo metti in Ritorno e il CRM glieli manda subito. Se il software non risponde, il CRM smette di riprovare fino al prossimo caricamento della pagina e te lo segnala nel riquadro." },
      { h: "Sicurezza", p: "La chiave permette solo di mandare task e di leggere i cambi di quelle task: non dà accesso al resto del CRM. Nel database ne resta solo un'impronta, per questo non si può rivedere. Metti in pausa per rifiutare temporaneamente le task in arrivo; Scollega elimina il collegamento ma lascia nel CRM le task già arrivate." },
    ],
  },
  {
    id: "pixel", t: "Pixel facile",
    intro: "Genera il codice degli eventi standard di Meta senza scriverlo a mano: scegli l'evento, compili i parametri che ti servono e copi il risultato.",
    blocchi: [
      { h: "Come funziona", p: "Primo riquadro: l'evento, diciassette in tutto, con il nome italiano e quello tecnico che Meta si aspetta. Sotto compare la descrizione di quando va usato. Secondo riquadro: i parametri, che cambiano con l'evento scelto e sono tutti facoltativi. Terzo riquadro: il codice, che si aggiorna mentre scrivi, con il pulsante per copiarlo." },
      { h: "I numeri all'italiana", p: "Nei campi di valore puoi scrivere come sei abituato — 1.500,00 oppure 3,50 — e il codice li converte nella forma che Meta richiede, 1500.00 e 3.50. Un valore scritto con la virgola e lasciato così sarebbe rifiutato senza che nulla te lo segnali." },
      { h: "Le scorciatoie dal cliente", p: "Entrando nella sezione da dentro un cliente, accanto all'evento compaiono i pulsanti per riempire i parametri con i suoi dati: la valuta, il valore medio dell'ordine, il settore come area tematica. Serve a non riscrivere ogni volta le stesse cose per lo stesso cliente." },
      { h: "Dove va incollato", p: "Nel punto del sito in cui l'evento avviene: nella pagina di ringraziamento per l'acquisizione contatto, in quella di conferma per l'acquisto. Deve stare dopo il codice base del pixel, altrimenti la funzione fbq non esiste ancora e l'evento non parte." },
      { h: "Perché conta per il resto del CRM", p: "L'evento che tracci qui è lo stesso che leggi nelle colonne Lead e Conversioni dell'Advertising, ed è la base del CPL e quindi del semaforo. Un evento piazzato nel posto sbagliato gonfia i numeri senza che nulla lo segnali: restano coerenti con sé stessi e sbagliati. Dopo l'installazione fai una conversione di prova e verifica che ne arrivi una sola." },
    ],
  },
  {
    id: "lingue", t: "La lingua dell'interfaccia",
    intro: "Il CRM parla italiano, inglese, spagnolo, francese e tedesco. Il selettore è in alto a destra, accanto ai pulsanti: un menu a tendina con i nomi delle lingue.",
    blocchi: [
      { h: "Dove si cambia", p: "In alto a destra c'è il blocchetto con l'etichetta della lingua — Lingua, Language, Idioma, Langue, Sprache — e il menu a tendina con i nomi delle lingue scritti nella loro lingua. L'etichetta stessa si traduce, così chi apre l'app in tedesco legge Sprache e capisce dove mettere le mani." },
      { h: "Cosa cambia", p: "Cambiano le voci di menu, i titoli delle sezioni, le etichette degli stati e delle fasi, i pulsanti ricorrenti e le intestazioni delle colonne. La scelta si salva con i dati dell'agenzia, quindi resta anche riaprendo l'app e viaggia nel backup." },
      { h: "Numeri, date e valuta", p: "Cambiando lingua cambiano anche i formati: in inglese 1.234,56 € diventa €1,234.56 e il 2 ottobre diventa Oct 02. Non è estetica — un anglofono che legge 1.234 capisce “uno virgola due”, cioè mille volte meno. Le date seguono la stessa regola." },
      { h: "Una valuta per cliente", p: "Nella scheda di ogni cliente, accanto all'account pubblicitario, c'è il campo Valuta: lasciandolo su “quella dell'agenzia” segue il valore predefinito, impostandone uno quel cliente usa il suo. Vale anche per i singoli sottoclienti, che possono differire dalla madre; se lasci il campo vuoto, il sottocliente segue la valuta della madre. È indipendente dalla lingua: puoi lavorare in italiano e seguire un cliente in dollari." },
      { h: "La valuta del compenso", p: "Sotto il compenso mensile c'è un campo a parte, Valuta del compenso, per il cliente che ha le campagne in una valuta ma ti paga in un'altra: per esempio campagne in USD e compenso in euro. Il compenso appare in quella valuta nella scheda, nel cruscotto del cliente, nell'elenco Clienti e in Economia. Lasciandolo vuoto, il compenso segue la valuta del cliente." },
      { h: "Dove si vede", p: "Entrando nel cliente tutti i suoi importi sono nella sua valuta, compresi avvisi, semaforo e grafici. Anche fuori dal cliente ogni importo che appartiene a lui resta nella sua valuta: le righe di Clienti, Sottoclienti, Panoramica, Advertising, Chiusure, Ciclo di vendita, Riepilogo mensile ed Economia, gli avvisi, il semaforo e il calendario. Quando una sezione mostra solo clienti con la stessa valuta, anche i suoi totali sono in quella valuta. Un sottocliente senza valuta propria usa quella del cliente madre." },
      { h: "Il limite dei totali d'agenzia", p: "Quando i clienti usano valute diverse, le somme a livello di agenzia le mettono insieme come se fossero la stessa cosa, e il numero che ne esce non corrisponde a niente. L'app non converte: servirebbe un tasso di cambio, e per farlo bene quello del giorno di ogni singola spesa. Per questo in Panoramica compare un avvertimento che elenca le valute in uso, e il consiglio di leggere gli importi dentro il singolo cliente. Le righe dei singoli clienti restano invece sempre nella loro valuta: solo i totali che li sommano usano quella dell'agenzia. I conteggi — campagne, lead, demo, chiusure — restano attendibili ovunque, perché non dipendono dalla valuta." },
      { h: "Le due lingue inglesi", p: "Inglese americano e inglese britannico sono voci distinte, perché differiscono in tutto ciò che conta qui: la data, il separatore delle migliaia e la valuta attesa — dollaro per il primo, sterlina per il secondo. Canada e Australia usano la voce americana scegliendo poi CAD o AUD, che si mostrano come CA$ e A$ per distinguerli dal dollaro statunitense. Le due varianti condividono lo stesso dizionario di parole." },
      { h: "Il simbolo lo sceglie la combinazione", p: "Il simbolo e la sua posizione non sono scritti da nessuna parte: nascono dall'incontro fra lingua e valuta. Lo stesso importo in euro è 1.234,56 € in italiano e €1,234.56 in inglese; in sterline diventa £1,234.56, in franchi 1.234,56 CHF." },
      { h: "Cambiando lingua", p: "Scegliendo una lingua la cui valuta abituale è diversa da quella in uso, l'app chiede se cambiare anche quella, senza farlo da sola. Il motivo è che non converte: cambierebbe il simbolo davanti a numeri rimasti identici, e mille euro diventerebbero mille dollari per finta. I clienti con una valuta propria restano come sono." },
      { h: "La valuta", p: "L\'agenzia ha una valuta predefinita nelle Impostazioni, e ogni cliente può averne una propria nella sua scheda: entrando in quel cliente, tutti i suoi importi si mostrano in quella valuta. Serve per i clienti che fatturano in franchi, sterline o dollari. Attenzione a una cosa: l'app non converte, mostra. Gli importi che inserisci devono essere già nella valuta del cliente." },
      { h: "Leggere i numeri in ingresso", p: "Nei campi e nelle importazioni puoi scrivere in qualunque convenzione: 1.234,56 e 1,234.56 vengono letti come lo stesso valore, perché l'app guarda quale separatore viene per ultimo — quello è sempre il decimale. Vale anche per i file esportati da piattaforme impostate in inglese." },
      { h: "Cosa resta in italiano, per ora", p: "I testi lunghi: le diagnosi degli avvisi, i consigli del semaforo e questa guida. Sono circa dodicimila parole, e tradurli bene è un lavoro a sé che si fa una lingua alla volta. Quando una frase non ha ancora traduzione compare in italiano, non vuota: una traduzione parziale non lascia buchi al posto delle parole." },
      { h: "Aggiungere una lingua", p: "Il meccanismo è un elenco di chiavi: aggiungere una lingua significa compilare quell'elenco, non rifare il software. Lo stesso vale per completare i testi lunghi, che si possono tradurre un capitolo alla volta senza toccare nient'altro." },
      { h: "Cosa non si traduce mai", p: "I tuoi dati: nomi di clienti, campagne, annunci, note del diario restano come li hai scritti. Si traduce il software, non il contenuto." },
    ],
  },
  {
    id: "dati", t: "Importazione e backup",
    intro: "Meno digitazione e più sicurezza: le due cose che fanno sopravvivere uno strumento come questo.",
    blocchi: [
      { h: "Importare", p: "Carichi il file esportato dalla piattaforma — Excel, CSV o TSV — oppure incolli le righe da un foglio. Se il file Excel ha più fogli, l'app apre quello con più righe e ti lascia cambiarlo da un menu. Riconosce le colonne più comuni, ti mostra l'anteprima e solo dopo scrive. Salva un profilo per ogni esportazione ricorrente: la volta dopo è un click." },
      { h: "Backup manuale", p: "È l'unico che ti protegge davvero. Scarica il file .json e tienilo in una cartella sincronizzata. Se non ne scarichi uno da più di sette giorni l'app te lo ricorda." },
      { h: "Copia automatica", p: "Una al giorno, ne conserva sei, ma resta dentro il browser: ti salva da un errore tuo o da dati rovinati, non dalla perdita del dispositivo. Nessuna pagina web può scrivere file sul tuo disco senza un tuo click." },
    ],
  },
  {
    id: "consulente", t: "Consulente",
    intro: "È una conversazione con un consulente che ha già davanti i tuoi numeri: non devi spiegargli il contesto ogni volta.",
    blocchi: [
      { h: "Cosa vede", p: "A ogni domanda l'app allega un riepilogo dei dati: campagne con le metriche calcolate, avvisi attivi, ad set principali, totali del ciclo di vendita con capacità e tempi di contatto, chiusure per funnel, passi dei funnel, task aperti e ultime annotazioni." },
      { h: "Ambito", p: "Dentro un cliente puoi scegliere se fargli vedere solo lui o tutto il portafoglio. Per un'analisi profonda tieni l'ambito ristretto: vede meno cose ma le vede tutte, perché il riepilogo è compatto e taglia a dodici campagne, sei ad set e otto task per cliente." },
      { h: "Analizza un report", p: "In cima alla sezione puoi allegare un file: Excel, CSV, PDF o anche una schermata. I fogli Excel vengono convertiti in tabella qui nel browser, tutti i fogli insieme. Poi estrae i numeri e ti mostra riga per riga cosa scriverebbe in Advertising, negli ad set, nel Ciclo di vendita e nelle Chiusure. Niente viene salvato finché non premi Inserisci nei dati." },
      { h: "Verifica sempre prima di confermare", p: "L'estrazione sbaglia più facilmente dalle schermate che dai CSV, e le colonne con nomi ambigui possono finire nel campo sbagliato. Guarda l'anteprima: due secondi lì valgono più di un'ora passata a capire perché un CPA non torna. Per volumi grandi resta più affidabile Importa dati, dove l'abbinamento delle colonne lo decidi tu." },
      { h: "Di cosa parlate", p: "Il primo menu sceglie l'ambito: l'agenzia nel suo insieme — con i numeri aggregati, il conteggio dei clienti, gli avvisi aperti e la distribuzione del semaforo — un cliente in particolare, scelto dal menu accanto senza doverci entrare, tutti i clienti uno per uno, oppure una domanda generica senza condividere nessun dato. L'ultima modalità serve quando vuoi un parere di metodo e non un'analisi." },
      { h: "Farsi creare le automazioni", p: "Puoi chiedere al consulente di tradurre una procedura in regole: per esempio le soglie del framework decisionale. Risponde spiegando, e sotto la risposta compare un riquadro con le regole proposte — innesco, cliente, azione — e il pulsante per crearle. Finché non lo premi non viene creato niente: la proposta resta una proposta." },
      { h: "Da cosa partire", p: "A chat vuota compaiono quattro spunti, diversi per ambito: sull'agenzia riguardano la distribuzione del tempo e i clienti che vanno peggio, sul cliente il messaggio e gli angoli, in modalità generica le domande di metodo. Cliccandone uno te lo scrive nel campo, dove puoi modificarlo prima di inviare." },
      { h: "Non lasciare le risposte nella chat", p: "Ogni risposta ha due pulsanti: Trasforma in task, che apre il modulo già compilato, e Salva nel diario, che la archivia come annotazione datata. Se un consiglio vale, deve diventare lavoro o memoria." },
      { h: "Fuori da Claude", p: "Dentro Claude non serve configurare nulla. Nella versione autonoma serve una chiave API personale, da incollare in fondo alla sezione: non metterla mai in una versione pubblicata online." },
    ],
  },
  {
    id: "file", t: "File e procedure",
    intro: "Due cose diverse nella stessa sezione: i collegamenti ai documenti che vivono altrove, e le procedure scritte qui dentro.",
    blocchi: [
      { h: "Collegamenti", p: "Strategie, brief, contratti e cartelle Drive restano dove sono: qui salvi il collegamento con categoria, cliente ed etichette, così accanto ai numeri del cliente trovi anche i suoi documenti." },
      { h: "Procedure", p: "Conviene scriverle qui: sono corte, cambiano spesso e finiscono nel backup insieme a tutto il resto. Ne trovi tre già pronte — rilevazione settimanale, chiusura del mese, primi sette giorni con un nuovo cliente. Se non le vedi, il riquadro in cima alla sezione le inserisce, oppure le prendi da Impostazioni → Contenuti pronti, dove ci sono anche le regole di automazione." },
      { h: "Materiale dell'agenzia", p: "Un file senza cliente vale per tutti: lo vedi anche quando sei dentro un singolo account, in fondo alla sezione." },
      { h: "I file KPI del cliente", p: "Dentro un cliente, in cima alla sezione File, c'è il riquadro Importa i numeri da un file KPI: carichi il foglio di riepilogo che il cliente tiene in Excel o CSV e i suoi numeri mese per mese vanno nel Ciclo di vendita, dopo un'anteprima e la tua conferma. Il file non resta salvato: si leggono i numeri e basta. Come funziona nel dettaglio lo trovi nel capitolo Ciclo di vendita e chiusure." },
      { h: "Cosa non fa", p: "Non ospita i file veri. Un PDF o un video dentro l'archivio del browser lo riempirebbero in fretta e renderebbero pesantissimo ogni backup: caricali su Drive e incolla qui il collegamento." },
    ],
  },
  {
    id: "leggere", t: "Leggere l'interfaccia",
    intro: "L'app è pensata per essere guardata ogni giorno su schermo: fondo scuro, un solo colore d'accento e i colori funzionali usati solo dove significano qualcosa.",
    blocchi: [
      { h: "I colori delle icone", p: "Ogni riquadro in cima alle sezioni ha un'icona colorata per tipo di numero. Rosso: soldi che escono e avvisi — spesa, budget, investito. Verde: soldi che entrano e ritorni — incassato, contratti, ROAS, chiusure. Ambra: costi unitari e tempo — CPL, CPA, costo per chiusura, ore, task. Blu: volumi e percentuali — lead, clienti, pipeline, tassi. Viola: attività pubblicitaria — CTR, clic, campagne, ritmo di erogazione." },
      { h: "Gli stati", p: "Le etichette seguono la stessa logica: verde per attivo, cliente, vinto e fatto; ambra per test, pausa e onboarding; bordeaux per critico, urgente e in ritardo. Il bordeaux dell'agenzia resta riservato a elementi attivi, pulsanti principali e avvisi critici: se vedi bordeaux, o è un comando o è un problema." },
      { h: "I tre livelli degli avvisi", p: "Bordo bordeaux per i critici, ambra per le attenzioni, grigio per le note. Le note non sono problemi — dicono per esempio che è troppo presto per giudicare una campagna — e infatti non entrano nei conteggi degli avvisi né nei numeri accanto alle voci di menu." },
      { h: "Tornare indietro", p: "La freccia in alto a sinistra, accanto al titolo, riporta sempre al cruscotto da qualunque sezione. Nella lista degli avvisi c'è anche un pulsante esplicito in cima." },
      { h: "Stampa e PDF", p: "Sullo schermo il tema è scuro, ma i report che generi in PDF e le stampe tornano su fondo bianco: nero pieno su carta sarebbe illeggibile e costoso da stampare." },
    ],
  },
  {
    id: "limiti", t: "Cosa non può fare",
    intro: "Meglio saperlo prima di costruirci sopra un processo. Non sono funzioni mancanti: sono il confine di un'app che vive nel browser.",
    blocchi: [
      { h: "Non è condivisa", p: "I dati stanno su questo dispositivo e in questo browser. Non c'è modo di far entrare un cliente a vedere solo il suo account, né di ritrovare gli stessi dati sul telefono." },
      { h: "Non riceve dati dall'esterno", p: "Niente lead dai moduli Meta in tempo reale, niente webhook in entrata, niente collegamento diretto alle API delle piattaforme: quelle richiedono chiavi di accesso custodite su un server." },
      { h: "Non lavora mentre dormi", p: "Automazioni e promemoria a orario funzionano solo con l'app aperta." },
      { h: "Quando serve il passo successivo", p: "Se una di queste tre cose diventa centrale nel tuo lavoro, la strada è ospitare l'applicazione su un backend. La struttura dei dati costruita qui non si butta: si riusa." },
    ],
  },
];

const GLOSSARIO = [
  ["CTR", "Percentuale di chi clicca sul totale di chi ha visto. Basso significa creatività o messaggio che non agganciano; fa salire anche il CPM."],
  ["CTR link", "Come il CTR ma contando solo i clic sul link, cioè chi è uscito davvero verso la tua pagina. È quello su cui va letta la soglia dello 0,80%: i clic totali comprendono anche like, commenti ed espansioni del testo, e gonfiano il dato."],
  ["CPC", "Quanto paghi un click. Se sale mentre il CPA sale, il problema è a monte: asta, pubblico, creatività."],
  ["CPC link", "Costo di ogni clic uscito verso la pagina. È il costo reale del traffico che ricevi."],
  ["CPL", "Costo per lead: spesa diviso lead. Nelle campagne a contatti è la metrica che porti al cliente, e la soglia di valutazione è il doppio del CPL obiettivo."],
  ["CPM", "Costo per mille visualizzazioni. Dipende dalla competizione dell'asta e dalla qualità percepita dell'annuncio."],
  ["CPA", "Costo per conversione. Se è alto ma il CPC è stabile, il problema è dopo il click: pagina, modulo, offerta."],
  ["ROAS", "Ricavi diviso spesa. Da solo non dice se guadagni: va confrontato con il ROAS di pareggio."],
  ["ROAS di pareggio", "Il ROAS sotto il quale il cliente perde soldi. Si ricava dal margine netto: con il 40% di margine il pareggio è a 2,5×."],
  ["Frequenza", "Quante volte la stessa persona ha visto l'annuncio. Alta significa pubblico troppo piccolo per il budget."],
  ["Show up", "Quanti si presentano davvero rispetto a quanti avevano prenotato. È il buco più costoso e il meno misurato."],
  ["Tasso di chiusura", "Chiusure sulle demo svolte. Va letto insieme allo show up: chiudere bene su poche demo non è un buon risultato."],
  ["Peso morto", "Un ad set o un annuncio che consuma budget senza portare risultati. Toglierlo non migliora il costo di chi resta, ma elimina la quota di spesa che la campagna paga a vuoto."],
  ["Attesa dopo modifica", "I tre giorni in cui il semaforo non giudica una campagna appena modificata: servono all'algoritmo per ritarare la consegna. Si avviano registrando l'intervento dalla scheda."],
  ["Semaforo", "Il giudizio sintetico su una campagna: rosso spegni, giallo non fare nulla, verde scala. Nasce da due indicatori soltanto, il rapporto con il KPI e il consumo del budget."],
  ["KPI", "Il costo obiettivo per risultato: il CPL massimo nelle campagne a contatti, il CPA massimo altrove. È il numero contro cui si misura il costo reale: senza, il semaforo resta grigio."],
  ["Scaling verticale e orizzontale", "Verticale è aumentare il budget della stessa campagna, e richiede che il budget pieno sia stato speso negli ultimi tre giorni. Orizzontale è allargarsi con nuovi pubblici o nuove campagne, ed è la strada quando la campagna non consuma quello che ha."],
  ["Errore pagamento", "Stato della campagna da usare quando si ferma per un problema di pagamento. Finché resta selezionato l'app segnala i giorni di fermo; togliendolo, scrive nel Diario quanto è durato."],
  ["Linea di business", "Etichetta che separa le campagne di attività diverse dentro lo stesso cliente. Campagne della stessa linea si sommano, di linee diverse no: medie e soglie perderebbero senso."],
  ["Data partenza", "Negli elenchi dei sottoclienti è il primo giorno di attività pubblicitaria, cioè l'inizio della campagna più vecchia. Se non ci sono campagne mostra la data da cui segui il cliente."],
  ["Finestra di analisi", "I sette giorni su cui si giudica una campagna già assestata. Serve perché la media dall'inizio si muove troppo lentamente per accorgersi di un peggioramento in corso."],
  ["Fasi Test, Matura, Attiva", "Test sono i primi tre giorni, in cui non si decide. Matura i sette successivi, in cui gli avvisi non diventano critici salvo quelli tecnici. Attiva dal decimo giorno, quando la valutazione si fa sulla finestra dei sette giorni."],
  ["Ritorno", "Valore dei contratti firmati diviso la spesa pubblicitaria. Diverso dal ROAS di piattaforma: è il numero che il cliente riconosce come suo."],
  ["Valore contratto", "Quanto è stato firmato. Diverso dall'incassato, che è quanto è entrato davvero."],
  ["Erogazione", "Quanto budget hai speso rispetto a quanto periodo è passato. Serve a non restare scoperto negli ultimi giorni."],
];

/* ================================================================== */
/* Importazione da file esportati dalle piattaforme                    */
/* ================================================================== */

/* Numeri scritti in mille modi diversi: 1.234,56 · 1,234.56 · € 1 234,56 */
const parseNumIntl = (s) => {
  if (typeof s === "number") return isFinite(s) ? s : 0;
  let t = String(s || "").replace(/[^\d,.\-]/g, "");
  if (!t) return 0;
  const lc = t.lastIndexOf(","), lp = t.lastIndexOf(".");
  if (lc > -1 && lp > -1) t = lc > lp ? t.replace(/\./g, "").replace(",", ".") : t.replace(/,/g, "");
  else if (lc > -1) t = /^-?\d{1,3}(,\d{3})+$/.test(t) ? t.replace(/,/g, "") : t.replace(",", ".");
  else if (lp > -1 && /^-?\d{1,3}(\.\d{3})+$/.test(t)) t = t.replace(/\./g, "");
  const n = parseFloat(t);
  return isFinite(n) ? n : 0;
};

/* Mese accettato come 2026-08, 08/2026, 1 ago 2026, 2026-08-14 */
const normalizzaMese = (s) => {
  const t = String(s || "").trim();
  let m = t.match(/^(\d{4})[-/](\d{1,2})/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, "0")}`;
  m = t.match(/^(\d{1,2})[-/](\d{4})$/);
  if (m) return `${m[2]}-${String(m[1]).padStart(2, "0")}`;
  m = t.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (m) return `${m[3]}-${String(m[2]).padStart(2, "0")}`;
  const d = new Date(t);
  return isNaN(d) ? "" : isoLocale(d).slice(0, 7);
};

const normalizzaData = (s, ordine = "") => {
  const t = String(s || "").trim();
  let m = t.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = t.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (m) {
    /* Di norma giorno/mese (Italia). Se il secondo numero supera 12 è per forza mese/giorno (formato USA),
       e lo stesso vale se l'app è impostata in inglese e la data è ambigua. */
    let g = +m[1], me = +m[2];
    if (me > 12 && g <= 12) [g, me] = [me, g];
    else if (g <= 12 && me <= 12 && (ordine === "md" || (!ordine && LINGUA === "en" && /\//.test(t)))) [g, me] = [me, g];
    if (me < 1 || me > 12 || g < 1 || g > 31) return "";
    return `${m[3]}-${String(me).padStart(2, "0")}-${String(g).padStart(2, "0")}`;
  }
  const d = new Date(t);
  return isNaN(d) ? "" : isoLocale(d);
};

/* Parser CSV/TSV con gestione delle virgolette */
const leggiTabella = (testo) => {
  const txt = String(testo || "").replace(/\r\n?/g, "\n").trim();
  if (!txt) return { intestazioni: [], righe: [] };
  const prima = txt.split("\n")[0];
  const sep = [";", "\t", ","].map((s) => [s, (prima.match(new RegExp(`\\${s}`, "g")) || []).length])
    .sort((a, b) => b[1] - a[1])[0][0];
  const out = []; let riga = [], campo = "", virg = false;
  for (let i = 0; i < txt.length; i++) {
    const c = txt[i];
    if (virg) {
      if (c === '"') { if (txt[i + 1] === '"') { campo += '"'; i++; } else virg = false; }
      else campo += c;
    } else if (c === '"') virg = true;
    else if (c === sep) { riga.push(campo.trim()); campo = ""; }
    else if (c === "\n") { riga.push(campo.trim()); out.push(riga); riga = []; campo = ""; }
    else campo += c;
  }
  riga.push(campo.trim()); out.push(riga);
  const intestazioni = (out.shift() || []).map((h) => h.replace(/^\uFEFF/, ""));
  const righe = out.filter((r) => r.some((c) => c !== ""));
  return { intestazioni, righe };
};

/* Excel: ogni foglio diventa una tabella in testo, così passa dagli stessi controlli del CSV */
const eExcel = (nome) => /\.(xlsx|xlsm|xls|ods)$/i.test(nome || "");

const fogliDaExcel = (buffer) => {
  const wb = XLSX.read(new Uint8Array(buffer), { type: "array", cellDates: true });
  return wb.SheetNames.map((n) => {
    const csv = XLSX.utils.sheet_to_csv(wb.Sheets[n], { FS: ";", dateNF: "yyyy-mm-dd", blankrows: false }).trim();
    return { nome: n, csv, righe: csv ? csv.split("\n").length - 1 : 0 };
  }).filter((f) => f.righe > 0);
};

const CAMPI_IMPORT = {
  clienti: {
    l: "Clienti e sottoclienti", chiave: "nome",
    campi: [
      { k: "nome", l: "Nome", alias: ["nome", "cliente", "nome cliente", "azienda", "name", "denominazione", "studio", "professionista"] },
      { k: "settore", l: "Settore", alias: ["settore", "categoria", "industry", "attivita", "attività"] },
      { k: "citta", l: "Sede", alias: ["citta", "città", "sede", "comune", "city"] },
      { k: "stato", l: "Stato", alias: ["stato", "status", "fase"] },
      { k: "fee", l: "Compenso mensile", num: true, alias: ["fee", "compenso", "canone", "retainer", "compenso mensile"] },
      { k: "budgetMese", l: "Budget adv mensile", num: true, alias: ["budget", "budget mensile", "budget adv", "spesa prevista"] },
      { k: "aov", l: "Valore medio ordine", num: true, alias: ["aov", "valore medio ordine", "scontrino medio", "ticket medio"] },
      { k: "margineLordo", l: "Margine lordo %", num: true, alias: ["margine", "margine lordo", "marginalita", "marginalità"] },
      { k: "referenteAmm", l: "Referente", alias: ["referente", "contatto", "titolare", "responsabile"] },
      { k: "emailAmm", l: "Email", alias: ["email", "e-mail", "mail", "posta"] },
      { k: "ragioneSociale", l: "Ragione sociale", alias: ["ragione sociale", "denominazione", "societa", "società"] },
      { k: "piva", l: "Partita IVA", alias: ["piva", "p iva", "partita iva", "vat"] },
      { k: "creato", l: "In gestione dal", data: true, alias: ["in gestione dal", "onboarding", "data onboarding", "attivo dal", "cliente dal", "data ingresso"] },
      { k: "dataContratto", l: "Data contratto", data: true, alias: ["data contratto", "contratto", "firma", "data firma", "firmato il", "contract date"] },
      { k: "accountAdv", l: "Account pubblicitario", alias: ["account", "account pubblicitario", "ad account", "account id", "act"] },
      { k: "cognome", l: "Cognome", alias: ["cognome", "surname", "last name"] },
      { k: "note", l: "Note", alias: ["note", "appunti", "commenti", "notes"] },
    ],
  },
  campagne: {
    l: "Campagne pubblicitarie", chiave: "nome",
    campi: [
      { k: "nome", l: "Nome campagna", alias: ["nome campagna", "nome della campagna", "campaign name", "campagna", "campaign", "nome"] },
      { k: "linea", l: "Linea di business", alias: ["linea", "linea di business", "business line", "reparto", "divisione", "brand"] },
      { k: "inizio", l: "Data di partenza", data: true, alias: ["data di partenza", "inizio", "data inizio", "start time", "starts", "partenza"] },
      { k: "fine", l: "Data di fine", data: true, alias: ["data di fine", "fine", "data fine", "stop time", "ends", "conclusa il"] },
      { k: "stato", l: "Stato", alias: ["stato", "status", "effective status", "delivery"] },
      { k: "piattaforma", l: "Piattaforma", alias: ["piattaforma", "platform", "canale"] },
      { k: "obiettivo", l: "Obiettivo", alias: ["obiettivo", "objective"] },
      { k: "budgetGiorno", l: "Budget giornaliero", num: true, alias: ["budget giornaliero", "daily budget", "budget al giorno"] },
      { k: "spesa", l: "Speso", num: true, alias: ["amount spent", "importo speso", "spesa", "costo", "cost", "spend"] },
      { k: "impression", l: "Impression", num: true, alias: ["impressions", "impressioni", "impression", "impr"] },
      { k: "click", l: "Clic totali", num: true, alias: ["clic totali", "clicks all", "clic tutti", "clicks", "clic", "click"] },
      { k: "clickLink", l: "Clic sul link", num: true, alias: ["link clicks", "clic sul link", "clic sul collegamento", "outbound clicks"] },
      { k: "lead", l: "Lead", num: true, alias: ["leads", "lead", "contatti", "moduli"] },
      { k: "conversioni", l: "Conversioni", num: true, alias: ["results", "risultati", "conversions", "conversioni", "acquisti", "purchases"] },
      { k: "ricavi", l: "Ricavi", num: true, alias: ["purchase conversion value", "valore di conversione", "conversion value", "revenue", "ricavi"] },
      { k: "valoreContratto", l: "Valore contratto", num: true, alias: ["contratto", "valore contratto", "contract value"] },
      { k: "incassato", l: "Incassato", num: true, alias: ["incassato", "cash collected", "pagato"] },
      { k: "frequenza", l: "Frequenza", num: true, alias: ["frequency", "frequenza"] },
      { k: "budget", l: "Budget", num: true, alias: ["budget"] },
    ],
  },
  creativi: {
    l: "Ad set e annunci", chiave: "nome",
    campi: [
      { k: "nome", l: "Nome ad set o annuncio", alias: ["nome", "nome ad set o annuncio", "nome annuncio", "nome creativo", "ad set name", "nome del gruppo di inserzioni", "ad name", "nome dell'inserzione", "adset", "annuncio"] },
      { k: "campagna", l: "Campagna di appartenenza", alias: ["campaign name", "nome della campagna", "campagna"] },
      { k: "adset", l: "Ad set di appartenenza", alias: ["ad set name", "nome ad set", "ad set", "adset", "gruppo di inserzioni"] },
      { k: "tipo", l: "Livello", alias: ["livello", "tipo", "level"] },
      { k: "spesa", l: "Speso", num: true, alias: ["amount spent", "importo speso", "spesa", "costo", "cost"] },
      { k: "impression", l: "Impression", num: true, alias: ["impressions", "impressioni"] },
      { k: "click", l: "Clic totali", num: true, alias: ["clic totali", "clicks all", "clic tutti", "clicks", "clic"] },
      { k: "clickLink", l: "Clic sul link", num: true, alias: ["link clicks", "clic sul link", "outbound clicks"] },
      { k: "lead", l: "Lead", num: true, alias: ["leads", "lead", "contatti", "moduli"] },
      { k: "conversioni", l: "Conversioni", num: true, alias: ["results", "risultati", "conversions", "conversioni"] },
      { k: "ricavi", l: "Ricavi", num: true, alias: ["conversion value", "valore di conversione", "revenue", "ricavi"] },
      { k: "frequenza", l: "Frequenza", num: true, alias: ["frequency", "frequenza"] },
    ],
  },
  fasi: {
    l: "Fasi delle campagne", chiave: "campagna",
    campi: [
      { k: "campagna", l: "Nome campagna", alias: ["campaign name", "campagna", "nome campagna", "campaign"] },
      { k: "tipo", l: "Fase", alias: ["fase", "phase", "tipo"] },
      { k: "inizio", l: "Dal", data: true, alias: ["dal", "date start", "reporting starts", "data inizio", "inizio"] },
      { k: "fine", l: "Al", data: true, alias: ["al", "date stop", "reporting ends", "data fine", "fine"] },
      { k: "spesa", l: "Speso", num: true, alias: ["amount spent", "importo speso", "spesa", "spend", "costo"] },
      { k: "impression", l: "Impression", num: true, alias: ["impressions", "impressioni", "impression", "impr"] },
      { k: "click", l: "Clic totali", num: true, alias: ["clic totali", "clicks all", "clic tutti", "clicks", "clic"] },
      { k: "clickLink", l: "Clic sul link", num: true, alias: ["link clicks", "clic sul link", "outbound clicks"] },
      { k: "lead", l: "Lead", num: true, alias: ["leads", "lead", "contatti"] },
      { k: "conversioni", l: "Conversioni", num: true, alias: ["results", "risultati", "conversions", "conversioni"] },
      { k: "ricavi", l: "Ricavi", num: true, alias: ["conversion value", "valore di conversione", "revenue", "ricavi"] },
      { k: "frequenza", l: "Frequenza", num: true, alias: ["frequency", "frequenza"] },
    ],
  },
  vendite: {
    l: "Ciclo di vendita", chiave: "periodo",
    campi: [
      { k: "periodo", l: "Mese", alias: ["mese", "month", "periodo", "data"] },
      { k: "lead", l: "Lead", num: true, alias: ["lead", "leads", "contatti"] },
      { k: "chiamate", l: "Chiamate", num: true, alias: ["chiamate", "calls", "dials"] },
      { k: "discoveryFissate", l: "Discovery fissate", num: true, alias: ["discovery fissate", "booked discovery", "appuntamenti fissati"] },
      { k: "discoverySvolte", l: "Discovery svolte", num: true, alias: ["discovery svolte", "show up discovery", "discovery held"] },
      { k: "demoFissate", l: "Demo fissate", num: true, alias: ["demo fissate", "booked demo"] },
      { k: "demoSvolte", l: "Demo svolte", num: true, alias: ["demo svolte", "demo held", "show up demo"] },
      { k: "offerte", l: "Offerte", num: true, alias: ["offerte", "proposals", "preventivi"] },
      { k: "spesa", l: "Spesa adv", num: true, alias: ["spesa", "spend", "ad spend", "costo"] },
      { k: "slot", l: "Slot disponibili", num: true, alias: ["slot", "capacita", "capacity"] },
      { k: "oreContatto", l: "Ore al primo contatto", num: true, alias: ["ore", "tempo", "response time", "primo contatto"] },
    ],
  },
  chiusure: {
    l: "Chiusure", chiave: "",
    campi: [
      { k: "nome", l: "Cliente chiuso", alias: ["nome", "name", "cliente", "contatto"] },
      { k: "azienda", l: "Azienda", alias: ["azienda", "company", "ragione sociale"] },
      { k: "email", l: "Email", alias: ["email", "e-mail", "mail"] },
      { k: "telefono", l: "Telefono", alias: ["telefono", "phone", "cellulare"] },
      { k: "data", l: "Data chiusura", alias: ["data", "date", "data chiusura", "closed"] },
      { k: "valoreContratto", l: "Valore contratto", num: true, alias: ["contratto", "valore", "amount", "deal value", "importo"] },
      { k: "incassato", l: "Incassato", num: true, alias: ["incassato", "cash", "paid", "pagato"] },
      { k: "campagna", l: "Campagna", alias: ["campagna", "campaign", "utm_campaign"] },
      { k: "adset", l: "Ad set", alias: ["adset", "ad set", "utm_content", "gruppo"] },
      { k: "annuncio", l: "Annuncio", alias: ["annuncio", "ad", "utm_term", "creativita"] },
      { k: "funnel", l: "Funnel", alias: ["funnel", "sorgente", "source", "utm_source"] },
      { k: "closer", l: "Closer", alias: ["closer", "venditore", "owner"] },
    ],
  },
};

/* Riconosce da sola le colonne più comuni di Meta, Google e dei fogli di calcolo */
const indovinaMappa = (tipo, intestazioni) => {
  const pulisci = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
  const teste = intestazioni.map((h) => ({ h, p: pulisci(h) }));
  const mappa = {};
  const presi = new Set();
  /* Prima gli abbinamenti esatti, poi quelli parziali: così una colonna che combacia
     con un campo non viene rubata da un altro che la contiene solo in parte. */
  CAMPI_IMPORT[tipo].campi.forEach((c) => {
    const esatto = teste.find((t) => !presi.has(t.h) && (t.p === pulisci(c.l) || c.alias.some((a) => t.p === pulisci(a))));
    if (esatto) { mappa[c.k] = esatto.h; presi.add(esatto.h); }
  });
  CAMPI_IMPORT[tipo].campi.forEach((c) => {
    if (mappa[c.k]) return;
    const parziale = teste.find((t) => !presi.has(t.h) && c.alias.some((a) => {
      const pa = pulisci(a);
      return pa.length >= 5 ? t.p.includes(pa) : ` ${t.p} `.includes(` ${pa} `);
    }));
    if (parziale) { mappa[c.k] = parziale.h; presi.add(parziale.h); }
  });
  return mappa;
};

/* ================================================================== */
/* Generatore della pagina: produce un file HTML autonomo              */
/* ================================================================== */

const paginaHTML = (f, cliente, agenzia, logo) => {
  const p = f.pagina || {};
  const punti = (p.punti || "").split("\n").map((x) => x.trim()).filter(Boolean);
  const esc = (s) => String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  /* Valori dentro <script>: JSON con "<" neutralizzato, così un "</script>" nel nome non rompe la pagina */
  const js = (v) => JSON.stringify(String(v || "")).replace(/</g, "\\u003c").replace(/\u2028/g, "\\u2028").replace(/\u2029/g, "\\u2029");
  const titolare = esc(cliente ? (cliente.ragioneSociale || cliente.nome) : agenzia.nome);
  return `<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${esc(p.titolo || f.nome)}</title>
<style>
:root{--nero:#0C0C0E;--acc:#8E2039;--paper:#F1EFEF;--line:#E0DBDC;--ink2:#565660}
*{box-sizing:border-box}
body{margin:0;background:var(--paper);color:#131316;font:16px/1.5 ui-sans-serif,system-ui,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;-webkit-font-smoothing:antialiased}
header{background:var(--nero);color:#fff;padding:18px 24px;display:flex;align-items:center;gap:12px}
header img{height:36px;width:auto;border-radius:50%}
header b{font-size:15px;font-weight:650}
main{max-width:860px;margin:0 auto;padding:48px 24px 72px}
h1{font-size:clamp(28px,5vw,44px);line-height:1.12;letter-spacing:-.02em;margin:0 0 16px}
.sub{font-size:clamp(16px,2.4vw,20px);color:var(--ink2);margin:0 0 32px;max-width:60ch}
ul{list-style:none;padding:0;margin:0 0 36px;display:grid;gap:12px}
li{position:relative;padding-left:30px;font-size:17px}
li::before{content:"";position:absolute;left:0;top:8px;width:12px;height:12px;border-radius:3px;background:var(--acc)}
form{background:#fff;border:1px solid var(--line);border-radius:12px;padding:24px;display:grid;gap:14px;max-width:460px}
label{font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:#8C8C96;display:block;margin-bottom:6px}
input{width:100%;font:inherit;padding:11px 12px;border:1px solid var(--line);border-radius:8px;background:#fff}
input:focus{outline:2px solid var(--acc);outline-offset:-1px}
button{background:var(--acc);color:#fff;border:0;border-radius:8px;padding:14px 18px;font:inherit;font-weight:600;cursor:pointer}
button:hover{background:#701829}
.consenso{display:flex;gap:10px;align-items:flex-start;font-size:13px;color:var(--ink2);line-height:1.45}
.consenso input{width:18px;height:18px;margin:2px 0 0;flex:none}
.ok{display:none;background:var(--nero);color:#fff;border-radius:12px;padding:24px;max-width:460px}
footer{border-top:1px solid var(--line);padding:22px 24px;color:var(--ink2);font-size:13px;text-align:center}
</style>
</head>
<body>
<header>${logo ? `<img src="${logo}" alt="">` : ""}<b>${esc(cliente ? cliente.nome : agenzia.nome)}</b></header>
<main>
  <h1>${esc(p.titolo || "")}</h1>
  <p class="sub">${esc(p.sottotitolo || "")}</p>
  ${punti.length ? `<ul>${punti.map((x) => `<li>${esc(x)}</li>`).join("")}</ul>` : ""}
  <form id="f">
    <div><label for="nome">Nome e cognome</label><input id="nome" name="nome" required></div>
    <div><label for="email">Email</label><input id="email" name="email" type="email" required></div>
    <div><label for="tel">Telefono</label><input id="tel" name="telefono"></div>
    <div class="consenso"><input id="consenso" type="checkbox" required><span>Acconsento al trattamento dei miei dati personali da parte di ${titolare} per essere ricontattato in merito a questa richiesta, ai sensi del Regolamento UE 2016/679 (GDPR).</span></div>
    <button type="submit">${esc(p.cta || "Invia la richiesta")}</button>
  </form>
  <div class="ok" id="ok"><b>Richiesta ricevuta.</b><p style="margin:8px 0 0">Ti ricontatteremo al più presto.</p></div>
</main>
<footer>${esc(cliente ? cliente.nome : agenzia.nome)} · pagina creata con ${esc(agenzia.nome)}</footer>
<script>
var HOOK = ${js(p.webhook)};
document.getElementById("f").addEventListener("submit", function(e){
  e.preventDefault();
  var d = {nome: nome.value, email: email.value, telefono: tel.value, funnel: ${js(f.nome)}, consenso: true, quando: new Date().toISOString()};
  if (HOOK) { fetch(HOOK, {method:"POST", mode:"no-cors", headers:{"Content-Type":"text/plain;charset=UTF-8"}, body: JSON.stringify(d)}).catch(function(){}); }
  else { console.log("Nessun webhook impostato:", d); }
  document.getElementById("f").style.display = "none";
  document.getElementById("ok").style.display = "block";
});
</script>
</body>
</html>`;
};

const scaricaFile = (nome, contenuto, tipo = "text/html") => {
  const a = document.createElement("a");
  const url = URL.createObjectURL(new Blob([contenuto], { type: tipo }));
  a.href = url; a.download = nome; document.body.appendChild(a); a.click();
  document.body.removeChild(a); setTimeout(() => URL.revokeObjectURL(url), 2000);
};

/* La finestra di analisi di una campagna matura: dalla rilevazione piu vicina
   a sette giorni fa fino a oggi. Senza rilevazioni non esiste, e si torna al totale. */
const finestraAnalisi = (data, k) => {
  const serie = (data.rilevazioni || []).filter((r) => r.campagnaId === k.id).sort((a, b) => (a.data < b.data ? -1 : 1));
  if (!serie.length) return null;
  const bersaglio = piu(-FINESTRA);
  const prima = serie.filter((r) => r.data <= bersaglio).slice(-1)[0] || serie[0];
  const iv = intervallo(k, prima);
  if (!iv) return null;
  const giorni = Math.max(1, Math.round((new Date(oggi()) - new Date(prima.data)) / 864e5));
  return { ...iv, giorni };
};

/* ==================================================================  */
/* ---------- Pixel facile: eventi standard di Meta ---------- */
const PARAM_PIXEL = {
  content_category: { l: "Area tematica", ph: "area_tematica", d: "Categoria della pagina o del prodotto." },
  content_ids: { l: "ID contenuto", ph: '["ABC123", "XYZ789"]', d: "ID del prodotto associato all'evento, per esempio lo SKU." },
  contents: { l: "Contenuto", ph: '[{"id": "ABC123", "quantity": 2}]', d: "Oggetti JSON con id e quantità: sono i due campi obbligatori." },
  currency: { l: "Valuta", d: "La valuta del valore indicato.", opts: true },
  value: { l: "Valore", ph: "3.00", d: "Valore economico del prodotto, o valore della persona che compie l'evento." },
  content_name: { l: "Nome contenuto", ph: "nome_contenuto", d: "Nome della pagina o del prodotto." },
  content_type: { l: "Tipo di contenuto", ph: "product", d: 'Vale "product" se gli ID sono di prodotti, "product_group" se sono di gruppi.' },
  status: { l: "Stato", ph: "stato", d: "Si usa con RegistrazioneCompletata per indicare lo stato della registrazione." },
  num_items: { l: "Numero di elementi", ph: "3", d: "Si usa con CheckoutIniziato: quanti articoli c'erano all'avvio." },
  search_string: { l: "Stringa di ricerca", ph: "stringa_ricerca", d: "Si usa con Ricerca: il testo cercato dalla persona." },
  predicted_ltv: { l: "Valore atteso nel tempo", ph: "50", d: "Durata prevista di un abbonato espressa come valore." },
};

const VALUTE_PIXEL = ["EUR", "USD", "GBP", "CHF", "CAD", "AUD", "BRL", "SEK", "NOK", "DKK", "PLN", "RON", "CZK", "HUF", "TRY", "AED", "JPY", "CNY", "INR", "MXN", "ARS", "CLP", "COP"];

const EVENTI_PIXEL = [
  { v: "Lead", l: "Acquisizione contatto", p: ["content_category", "content_name", "currency", "value"],
    d: "Quando una registrazione è completata e vengono inserite le informazioni di contatto: email, telefono e simili." },
  { v: "CompleteRegistration", l: "Registrazione completata", p: ["content_name", "currency", "status", "value"],
    d: "Quando un modulo di registrazione o di iscrizione viene inviato completo." },
  { v: "Schedule", l: "Appuntamento fissato", p: [],
    d: "Quando una persona fissa un appuntamento scegliendo data e ora." },
  { v: "Contact", l: "Contatto", p: [],
    d: "Quando una persona avvia un contatto: telefonata, messaggio, email, chat." },
  { v: "SubmitApplication", l: "Invio modulo", p: [],
    d: "Quando una persona invia una candidatura o una richiesta per un servizio o un programma." },
  { v: "Purchase", l: "Acquisto", p: ["content_ids", "content_name", "content_type", "contents", "currency", "num_items", "value"],
    d: "Quando un acquisto viene concluso e la persona arriva sulla pagina di conferma." },
  { v: "InitiateCheckout", l: "Checkout iniziato", p: ["content_category", "content_ids", "contents", "currency", "num_items", "value"],
    d: "Quando una persona avvia il pagamento senza averlo completato." },
  { v: "AddPaymentInfo", l: "Aggiunta info pagamento", p: ["content_category", "content_ids", "contents", "currency", "value"],
    d: "Quando i dati di pagamento vengono salvati durante il checkout." },
  { v: "AddToCart", l: "Aggiunta al carrello", p: ["content_ids", "content_name", "content_type", "contents", "currency", "value"],
    d: "Quando un prodotto viene messo nel carrello." },
  { v: "AddToWishlist", l: "Aggiunta alla lista desideri", p: ["content_ids", "content_name", "content_category", "contents", "currency", "value"],
    d: "Quando un prodotto viene salvato nella lista dei desideri." },
  { v: "ViewContent", l: "Visualizzazione contenuto", p: ["content_ids", "content_category", "content_name", "content_type", "contents", "currency", "value"],
    d: "Una visita a una pagina che ti interessa. Dice che qualcuno è arrivato lì, non cosa ci ha fatto." },
  { v: "Search", l: "Ricerca", p: ["content_category", "content_ids", "contents", "currency", "search_string", "value"],
    d: "Quando una persona cerca qualcosa sul sito." },
  { v: "StartTrial", l: "Inizio prova", p: ["currency", "predicted_ltv", "value"],
    d: "Quando una persona avvia una prova gratuita del prodotto o del servizio." },
  { v: "Subscribe", l: "Iscrizione", p: ["currency", "predicted_ltv", "value"],
    d: "Quando una persona attiva un abbonamento a pagamento." },
  { v: "Donate", l: "Donazione", p: [],
    d: "Quando una persona fa una donazione alla tua organizzazione o alla tua causa." },
  { v: "CustomizeProduct", l: "Prodotto personalizzato", p: [],
    d: "Quando una persona personalizza un prodotto, per esempio scegliendone il colore." },
  { v: "FindLocation", l: "Ricerca punto vendita", p: [],
    d: "Quando una persona cerca dove sei, con l'intenzione di venirci." },
];

/* Sistema a semaforo                                                  */
/* Due soli indicatori — rapporto con il KPI e ritmo di spesa — e tre
   segnali. Le fasi sono le stesse del resto dell'app: Test è il Proof
   of Concept, Matura è il Validation Period, Attiva è Scaling.        */

const SEGNALI = [
  { v: "verde", l: "Scala", c: "#3FA45C" },
  { v: "giallo", l: "Non fare nulla", c: "#C99A3A" },
  { v: "rosso", l: "Spegni", c: "#D8455F" },
  { v: "grigio", l: "Non ancora valutabile", c: "#6C6C79" },
];

/* I dati su cui si giudica: la finestra recente se c'è, altrimenti il totale */
const datiSemaforo = (data, k) => {
  const d = calc(k);
  const F = faseCampagna(k);
  if (F.fase === "attiva") {
    const w = finestraAnalisi(data, k);
    if (w) return { d: { ...w, ld: w.ld || 0 }, dove: `ultimi ${w.giorni} giorni`, finestra: true };
  }
  return { d, dove: "dall'inizio", finestra: false };
};

/* ---------- Analisi 30 / 7 / 3 ----------
   Tre finestre mobili che finiscono ieri: il giorno in corso è ancora parziale e
   confrontarlo con giorni interi falserebbe la lettura. I numeri si ricostruiscono
   dalle finestre compilate in proporzione ai giorni che cadono dentro l'intervallo,
   quindi una settimana che sporge contribuisce solo per la sua parte. Se la campagna
   è partita dopo l'inizio teorico, la finestra parte dalla sua data di partenza. */
const ORIZZONTI = [{ n: 30, l: "Ultimi 30 giorni" }, { n: 7, l: "Ultimi 7 giorni" }, { n: 3, l: "Ultimi 3 giorni" }];

const finestraMobile = (data, k, n) => {
  const al = sommaGiorni(oggi(), -1);
  let dal = sommaGiorni(al, -(n - 1));
  let tagliata = false;
  if (k.inizio && k.inizio > dal) { dal = k.inizio; tagliata = true; }
  if (dal > al) return null;
  const giorni = giorniTra(dal, al);
  const CH = ["spesa", "impression", "click", "clickLink", "lead", "conversioni", "ricavi"];
  const g = { fq: 0 };
  CH.forEach((c) => { g[c] = 0; });
  let nFin = 0, parziali = 0, coperti = 0;
  (data.periodi || []).filter((pz) => pz.campagnaId === k.id && pz.inizio && pz.fine).forEach((pz) => {
    const a = pz.inizio > dal ? pz.inizio : dal;
    const b = pz.fine < al ? pz.fine : al;
    if (a > b) return;
    const durata = giorniTra(pz.inizio, pz.fine);
    const dentro = giorniTra(a, b);
    if (durata <= 0 || dentro <= 0) return;
    const q = Math.min(1, dentro / durata);
    nFin += 1;
    if (q < 1) parziali += 1;
    coperti += dentro;
    CH.forEach((c) => { g[c] += (pz[c] || 0) * q; });
    g.fq += (pz.frequenza || 0) * (pz.impression || 0) * q;
  });
  const sp = g.spesa, imp = g.impression, cl = g.click, clk = g.clickLink, cv = g.conversioni, rc = g.ricavi;
  const ld = g.lead || (k.obiettivo === "lead" ? cv : 0);
  return {
    n, dal, al, giorni, tagliata, nFin, parziali, coperti, vuota: nFin === 0 || sp <= 0,
    sp, imp, cl, clk, ld, cv, rc,
    ctr: imp ? (cl / imp) * 100 : 0,
    ctrLink: imp ? ((clk || cl) / imp) * 100 : 0,
    cpcLink: (clk || cl) ? sp / (clk || cl) : 0,
    cpm: imp ? (sp / imp) * 1000 : 0,
    cpl: ld ? sp / ld : 0,
    cpa: cv ? sp / cv : 0,
    roas: sp ? rc / sp : 0,
    freq: imp ? g.fq / imp : 0,
  };
};

/* Lo stesso giudizio del semaforo, applicato a una finestra di date invece che
   alla fase: costo contro KPI e ritmo di spesa, nient'altro. */
const semaforoFinestra = (k, kpi, w) => {
  const seg = (v) => SEGNALI.find((s) => s.v === v) || {};
  const out = (v, perche, extra = {}) => ({ segnale: v, colore: seg(v).c, etichetta: seg(v).l, perche, ...extra });
  if (!w || w.vuota) return out("grigio", "Nessuna spesa registrata in questa finestra.", { costo: 0, rapporto: 0, quota: null });
  const conv = w.ld > 0 ? w.ld : w.cv;
  const costo = conv > 0 ? w.sp / conv : 0;
  const rapporto = kpi && costo ? (costo / kpi) * 100 : 0;
  const bg = budgetDelGiorno(k, w.al);
  const grezzo = bg > 0 ? (w.sp / w.giorni) / bg : null;
  /* Sotto un decimo del previsto il budget in scheda è sbagliato, non la campagna */
  const quota = grezzo === null || grezzo < 0.1 ? null : grezzo;
  const base = { costo, rapporto, quota, conv };
  if (!kpi) return out("grigio", "Manca il KPI: niente con cui confrontare il costo.", base);
  if (w.sp < kpi) return out("grigio", `Ha speso ${eur(w.sp)}, meno di una volta il KPI: troppo poco per leggerla.`, base);
  if (conv === 0) return out("rosso", `Nessun risultato con ${eur(w.sp)} spesi.`, base);
  if (rapporto > 150) return out("rosso", `${fmt(rapporto, 0)}% del KPI.`, base);
  if (rapporto > 100) return out("giallo", `${fmt(rapporto, 0)}% del KPI: sopra obiettivo, dentro la tolleranza.`, base);
  if (quota !== null && quota < 0.9) return out("giallo", `Dentro il KPI ma eroga il ${fmt(quota * 100, 0)}% del budget giornaliero.`, base);
  return out("verde", `${fmt(rapporto, 0)}% del KPI, budget consumato.`, base);
};

/* Una campagna appena modificata è in valutazione: non si tocca e non si giudica.
   Serve a tutte le viste, non solo al semaforo. */
const inValutazione = (k) => {
  if (!k.ultimaModifica) return null;
  const g = giorniTra(k.ultimaModifica, oggi());
  if (g > ATTESA_MODIFICA) return null;
  return { dal: k.ultimaModifica, fino: sommaGiorni(k.ultimaModifica, ATTESA_MODIFICA),
    passati: g - 1, restano: ATTESA_MODIFICA - (g - 1), nota: k.notaModifica || "" };
};

/* Scaduti i tre giorni la campagna torna giudicabile: finché non la rivedi
   resta una campanella accesa, perché è il momento in cui la decisione va presa. */
const valutazionePronta = (k) => {
  if (!k.ultimaModifica || k.modificaVista === "si") return null;
  const g = giorniTra(k.ultimaModifica, oggi());
  if (g <= ATTESA_MODIFICA) return null;
  return { dal: k.ultimaModifica, scaduta: sommaGiorni(k.ultimaModifica, ATTESA_MODIFICA),
    da: g - ATTESA_MODIFICA - 1, nota: k.notaModifica || "" };
};

const Campanella = ({ k, size = 13 }) => {
  const p = valutazionePronta(k);
  if (!p) return null;
  return (
    <span title={`Valutazione pronta dal ${dataItA(p.scaduta)}${p.nota ? ` — ${p.nota}` : ""}`}
      style={{ color: "#C99A3A", display: "inline-flex", verticalAlign: "middle" }}>
      <BellRing size={size} strokeWidth={2} aria-label="Valutazione pronta" />
    </span>
  );
};

/* Il giudizio scrive importi nei suoi testi: sempre nella valuta del cliente della campagna */
const semaforo = (data, k) => conValuta(valutaCliente(data, k.clienteId), () => semaforoBase(data, k));
const semaforoBase = (data, k) => {
  const b = k.bmk || {};
  const F = faseCampagna(k);
  const { d: v, dove } = datiSemaforo(data, k);
  const tot = calc(k);
  /* Il KPI: nelle campagne a contatti è il CPL, il primo punto di conversione.
     Senza CPL impostato vale il CPA, così niente resta grigio senza motivo. */
  const suLead = suiLead(k, tot);
  const kpi = kpiCosto(k, tot);
  const kpiNome = suLead ? "CPL" : "CPA";
  const E = erogazioneAttesa(data, k);

  const esito = (col, perche, azione, extra = {}) => ({
    colore: (SEGNALI.find((s) => s.v === col) || {}).c,
    segnale: col, etichetta: (SEGNALI.find((s) => s.v === col) || {}).l,
    fase: F.fase, faseL: { test: "Proof of Concept", matura: "Validation", attiva: "Scaling e mantenimento" }[F.fase] || "—",
    giorni: F.g, dove, perche, azione, kpi, kpiNome, ...extra,
  });

  if (k.archiviata === "si" || ["conclusa", "pausa"].includes(k.stato)) {
    return esito("grigio", "La campagna non è in corso.", "Niente da decidere.");
  }
  if (k.stato === "errore") {
    return k.causaErrore && k.causaErrore !== "pagamento"
      ? esito("rosso", `Ferma: ${causaFermo(k).l.toLowerCase()}.`, `${causaFermo(k).azione} Finché è ferma, nessun dato è attendibile.`)
      : esito("rosso", "Ferma per errore di pagamento.", "Sistema il metodo di pagamento: finché è ferma, nessun dato è attendibile.");
  }
  if (!kpi) {
    return esito("grigio", "Manca il KPI: senza costo obiettivo non c'è niente con cui confrontare il costo reale.",
      suLead
        ? "Apri la campagna e imposta il CPL massimo nei benchmark: nelle campagne a contatti è il numero su cui si regge tutto il sistema."
        : "Apri la campagna e imposta il CPA massimo: è il numero su cui si regge tutto il sistema.");
  }

  /* Dopo una modifica il giudizio si sospende per tre giorni */
  if (k.ultimaModifica && giorniTra(k.ultimaModifica, oggi()) <= ATTESA_MODIFICA) {
    const passati = giorniTra(k.ultimaModifica, oggi()) - 1;
    const quando = sommaGiorni(k.ultimaModifica, ATTESA_MODIFICA);
    return esito("grigio",
      `In valutazione dopo modifica: cambiata il ${dataItA(k.ultimaModifica)}${k.notaModifica ? ` — ${k.notaModifica}` : ""}, sono passati ${int(passati)} ${passati === 1 ? "giorno" : "giorni"} sui tre.`,
      `Non toccarla e non giudicarla fino al ${dataItA(quando)}. Dopo un intervento l'algoritmo ritara la consegna, e i numeri dei primi giorni descrivono quella transizione, non la campagna.`,
      { attesa: true, riprende: quando, etichettaBreve: "In valutazione" });
  }

  /* Il costo reale: per lead se ci sono lead, altrimenti per conversione */
  const conv = v.ld > 0 ? v.ld : v.cv;
  const costo = conv > 0 ? v.sp / conv : 0;
  const rapporto = costo ? (costo / kpi) * 100 : 0;

  /* Ritmo di spesa: il budget giornaliero viene consumato? */
  /* Confronto fra due grandezze omogenee: euro spesi in un giorno contro euro previsti in un giorno */
  const attesoGiorno = E && E.giorni && E.atteso > 0 && !E.sospetto ? E.atteso / E.giorni : null;
  const spesaGiorno = v.giorni ? v.sp / v.giorni : (E && E.giorni ? v.sp / E.giorni : 0);
  const grezzo = attesoGiorno && spesaGiorno ? spesaGiorno / attesoGiorno : null;
  /* Sotto un decimo del previsto non è sotto-erogazione: è un budget scritto male.
     Meglio non dare un giudizio che darne uno fondato su un numero sbagliato. */
  const pieno = grezzo === null || grezzo < 0.1 ? null : grezzo;
  const pacing = pieno === null ? null : pieno >= 0.9;

  /* Checkpoint di fase: prima non si giudica */
  if (F.fase === "test") {
    if (tot.sp < kpi) {
      return esito("grigio", `Ha speso ${eur(tot.sp)} contro un minimo di ${eur(kpi)}, pari a una volta il KPI.`,
        `Aspetta: il primo giudizio si dà a ${eur(kpi)} di spesa, quello solido a ${eur(kpi * 2)}.`,
        { costo, rapporto, pacing, checkpoint: `spesa ${eur(tot.sp)} di ${eur(kpi)}` });
    }
  }
  if (F.fase === "matura" && F.g < TEST_2) {
    return esito("grigio", `È in Validation da ${F.g} giorni: il periodo di attesa è di 7 giorni e non è ancora concluso.`,
      "Non toccarla. Il giudizio arriva al decimo giorno di vita della campagna.",
      { costo, rapporto, pacing, checkpoint: `giorno ${F.g} di ${TEST_2}` });
  }

  /* Rosso: nessuna conversione, oppure costo oltre il 150% del KPI */
  if (conv === 0) {
    return esito("rosso", `Nessuna conversione ${dove}, con ${eur(v.sp)} spesi.`,
      "Spegnila. Prima però verifica con una conversione di prova che il tracciamento registri: zero può voler dire rotto, non inefficace.",
      { costo, rapporto, pacing });
  }
  if (rapporto > 150) {
    return esito("rosso", `${kpiNome} reale ${eur(costo)} contro un obiettivo di ${eur(kpi)}: ${fmt(rapporto, 0)}% del target.`,
      "Spegnila e scrivi nel Diario quale angolo ha fallito, per non ritestarlo sullo stesso pubblico.",
      { costo, rapporto, pacing });
  }
  if (rapporto > 100) {
    return esito("giallo", `${kpiNome} reale ${eur(costo)} contro un obiettivo di ${eur(kpi)}: ${fmt(rapporto, 0)}% del target, dentro la fascia di tolleranza.`,
      "Non fare nulla. Non spegnere e non aumentare il budget: sopra il KPI ma sotto il 150% si lascia lavorare.",
      { costo, rapporto, pacing });
  }

  /* Dentro il KPI: il verde dipende dal ritmo di spesa */
  if (pacing === false) {
    return esito("giallo", `Costo ${eur(costo)}, dentro il KPI, ma non consuma tutto il budget giornaliero.`,
      "Niente aumenti di budget: su una campagna che già non eroga non funzionerebbero. Semmai allarga in orizzontale, con nuovi pubblici o nuove campagne.",
      { costo, rapporto, pacing });
  }
  if (pacing === null) {
    return esito("giallo", `Costo ${eur(costo)}, dentro il KPI, ma non so se consuma il budget: manca il budget giornaliero o le date.`,
      "Compila budget giornaliero e date nella scheda: senza, il verde non si può dare.",
      { costo, rapporto, pacing });
  }
  return esito("verde", `Costo ${eur(costo)} contro KPI ${eur(kpi)}: ${fmt(rapporto, 0)}% del target, e il budget viene consumato.`,
    F.fase === "attiva"
      ? "Scala. Per l'aumento di budget sulla stessa campagna serve il budget pieno speso negli ultimi 3 giorni consecutivi: se non ce l'hai, allarga in orizzontale."
      : "Promuovila alla fase successiva e continua a osservarla.",
    { costo, rapporto, pacing });
};

/* ================================================================== */
/* Automazioni: valutazione delle regole                               */
/* ================================================================== */

/* Ogni evento ha una chiave: se è già nell'elenco degli eseguiti non riparte. */
/* Campagne con i numeri "veri": totali dalle fasi quando ci sono, data di partenza reale,
   indicatori e avvisi. Usata sia dalle schermate sia dalle automazioni, così le due cose
   giudicano sempre gli stessi numeri. Esclude i clienti archiviati; le campagne archiviate
   restano (senza avvisi) perché le schermate le mostrano a parte. */
/* Redditività del cliente dal Ciclo di vendita: contratti (o, se mancano, incassato) sulla spesa
   degli ultimi tre mesi che hanno numeri commerciali. La spesa è quella del periodo; se non è scritta,
   quella delle finestre delle sue campagne in quei mesi. In attivo se supera il pareggio (o 1×).
   Un sottocliente senza ciclo suo eredita il giudizio del cliente madre. */
const redditivitaClienti = (data) => {
  const out = {};
  const spesaFin = {};
  const cliDiCamp = new Map(data.campagne.map((k) => [k.id, k.clienteId]));
  (data.periodi || []).forEach((pz) => {
    const cid = cliDiCamp.get(pz.campagnaId); if (!cid || !pz.inizio) return;
    const key = `${cid}|${pz.inizio.slice(0, 7)}`; spesaFin[key] = (spesaFin[key] || 0) + (pz.spesa || 0);
  });
  const perCli = {};
  (data.vendite || []).forEach((v) => { (perCli[v.clienteId] = perCli[v.clienteId] || []).push(v); });
  Object.entries(perCli).forEach(([cid, l]) => {
    const mesi = l.filter((v) => (v.valoreContratto || 0) + (v.incassato || 0) > 0).sort((a, b) => (a.periodo < b.periodo ? 1 : -1)).slice(0, 3);
    if (!mesi.length) return;
    const ct = mesi.reduce((x, v) => x + ((v.valoreContratto || 0) || (v.incassato || 0)), 0);
    const sp = mesi.reduce((x, v) => x + ((v.spesa || 0) || spesaFin[`${cid}|${v.periodo}`] || 0), 0);
    if (!sp) return;
    const z = data.clienti.find((x) => x.id === cid);
    const soglia = Math.max(1, roasPareggio(z));
    out[cid] = { ritorno: ct / sp, mesi: mesi.length, soglia, attivo: ct / sp >= soglia };
  });
  data.clienti.forEach((z) => { if (!out[z.id] && z.padreId && out[z.padreId] && out[z.padreId].attivo && !perCli[z.id]) out[z.id] = out[z.padreId]; });
  return out;
};

const campagneCalcolate = (data) => {
  const redd = redditivitaClienti(data);
  const perCli = {};
  data.campagne.forEach((k) => {
    if (!["attiva", "test"].includes(k.stato)) return;
    const g = perCli[k.clienteId] || (perCli[k.clienteId] = { sp: 0, imp: 0 });
    g.sp += k.spesa || 0; g.imp += k.impression || 0;
  });
  const cpmDi = (id) => (perCli[id] && perCli[id].imp ? (perCli[id].sp / perCli[id].imp) * 1000 : 0);
  /* Se una campagna ha fasi registrate, i suoi totali nascono da quelle */
  const perCamp = {};
  (data.periodi || []).forEach((pz) => {
    const g = perCamp[pz.campagnaId] || (perCamp[pz.campagnaId] = { spesa: 0, impression: 0, click: 0, clickLink: 0, lead: 0, conversioni: 0, ricavi: 0, fq: 0, n: 0 });
    g.spesa += pz.spesa || 0; g.impression += pz.impression || 0; g.click += pz.click || 0;
    g.clickLink += pz.clickLink || 0; g.lead += pz.lead || 0;
    g.conversioni += pz.conversioni || 0; g.ricavi += pz.ricavi || 0;
    g.fq += (pz.frequenza || 0) * (pz.impression || 0); g.n += 1;
  });
  const conFasi = (k) => {
    const g = perCamp[k.id];
    /* Fasi generate ma ancora vuote non devono azzerare i dati scritti a mano */
    if (!g || (g.spesa + g.impression + g.click + g.conversioni + g.ricavi) === 0) return k;
    /* Conversioni e ricavi sono l'esito commerciale, non un dato della piattaforma:
       se li hai scritti tu restano tuoi, le finestre non li sovrascrivono. */
    const mio = k.esitoManuale === "si";
    return { ...k, spesa: g.spesa, impression: g.impression, click: g.click, clickLink: g.clickLink, lead: g.lead,
      conversioni: mio ? (k.conversioni || 0) : g.conversioni,
      ricavi: mio ? (k.ricavi || 0) : g.ricavi,
      frequenza: g.impression ? g.fq / g.impression : (k.frequenza || 0), daFasi: g.n };
  };
  /* La data di partenza usata nei calcoli è quella reale: se le finestre mostrano spesa
     prima della data in scheda, vince la spesa. "inizioMeta" conserva quella scritta. */
  const conInizio = (k) => { const vero = inizioReale(data, k); return vero === k.inizio ? k : { ...k, inizio: vero, inizioMeta: k.inizio }; };
  const idArch = new Set(data.clienti.filter((z) => z.archiviato === "si").map((z) => z.id));
  /* I testi degli avvisi contengono importi: si scrivono nella valuta del cliente della campagna,
     qualunque sia il cliente aperto quando vengono calcolati. */
  return data.campagne.filter((k) => !idArch.has(k.clienteId)).map(conFasi).map(conInizio).map((k) => conValuta(valutaCliente(data, k.clienteId), () => ({ ...k, d: calc(k),
    al: k.archiviata === "si" ? [] : avvisi(k, {
      cliente: data.clienti.find((z) => z.id === k.clienteId),
      prec: ultimaRilevazione(data, k),
      cpmMedio: cpmDi(k.clienteId),
      serie: (data.rilevazioni || []).filter((r) => r.campagnaId === k.id).sort((a, b) => (a.data < b.data ? -1 : 1)),
      finestra: finestraAnalisi(data, k),
      erogazione: erogazioneAttesa(data, k),
      redditivita: redd[k.clienteId] && redd[k.clienteId].attivo ? redd[k.clienteId] : null,
    }) })));
};

function valutaAutomazioni(data, campagne) {
  const out = [];
  const g = oggi(), mm = mese(g);
  const fatti = new Set(data.eseguiti || []);
  const nomeC = (id) => (data.clienti.find((z) => z.id === id) || {}).nome || "Agenzia";

  (data.automazioni || []).filter((a) => a.attiva === "si").forEach((a) => {
    const suo = (id) => !a.clienteId || a.clienteId === id;
    /* rif = la chiave senza data: serve a non riaprire ogni giorno un task identico
       a uno ancora aperto. clienteId fa nascere il task sul cliente giusto. */
    const push = (chiave, testo, dettaglio, contesto, clienteId = "") => {
      const rif = chiave.replace(/\|\d{4}-\d{2}(-\d{2})?$/, "");
      if (!fatti.has(chiave)) out.push({ chiave, rif, aut: a, testo, dettaglio, contesto, clienteId });
    };

    if (a.trigger === "avviso") {
      campagne.filter((k) => suo(k.clienteId)).forEach((k) => k.al.forEach((al) => {
        if (a.filtroAvviso && al.metrica !== a.filtroAvviso) return;
        if (al.gravita === "nota") return; /* le note informative non sono problemi da girare in task */
        if (a.soloCritici === "si" && al.gravita !== "critica") return;
        push(`${a.id}|${k.id}|${al.metrica}|${g}`, `${al.titolo} · ${k.nome}`, al.azione,
          { cliente: nomeC(k.clienteId), campagna: k.nome, metrica: al.metrica, valore: al.valore, soglia: al.soglia, gravita: al.gravita }, k.clienteId);
      }));
    }

    if (a.trigger === "metrica") {
      campagne.filter((k) => suo(k.clienteId) && ["attiva", "test"].includes(k.stato)).forEach((k) => {
        const v = k.d[a.metrica] || 0;
        const ok = a.operatore === "gt" ? v > a.valore : v < a.valore;
        if (!ok) return;
        const et2 = et(METRICHE_TRIGGER, a.metrica);
        push(`${a.id}|${k.id}|${g}`, `${et2} ${a.operatore === "gt" ? "sopra" : "sotto"} soglia · ${k.nome}`,
          `Valore rilevato ${fmt(v)} contro una soglia di ${fmt(a.valore)}.`,
          { cliente: nomeC(k.clienteId), campagna: k.nome, metrica: et2, valore: v, soglia: a.valore }, k.clienteId);
      });
    }

    if (a.trigger === "budget") {
      campagne.filter((k) => suo(k.clienteId) && k.stato === "attiva" && k.budget > 0).forEach((k) => {
        const q = k.d.usato * 100;
        if (q < a.valore) return;
        push(`${a.id}|${k.id}|${g}`, `Budget al ${fmt(q, 1)}% · ${k.nome}`,
          conValuta(valutaCliente(data, k.clienteId), () => `Speso ${eur(k.d.sp)} su un budget di ${eur(k.budget)}.`),
          { cliente: nomeC(k.clienteId), campagna: k.nome, quota: q, spesa: k.d.sp, budget: k.budget }, k.clienteId);
      });
    }

    if (a.trigger === "task") {
      (data.task || []).filter((t) => t.stato !== "fatto" && t.scadenza && t.scadenza < g && suo(t.clienteId)).forEach((t) => {
        push(`${a.id}|${t.id}|${g}`, `Task in ritardo: ${t.titolo}`, `Scadeva il ${dataIt(t.scadenza)}.`,
          { cliente: nomeC(t.clienteId), task: t.titolo, scadenza: t.scadenza }, t.clienteId);
      });
    }

    if (a.trigger === "chiusura") {
      (data.chiusure || []).filter((c) => suo(c.clienteId)).forEach((c) => {
        push(`${a.id}|${c.id}`, `Chiusura registrata: ${c.nome}`,
          conValuta(valutaCliente(data, c.clienteId), () => `${eur(c.valoreContratto)} di contratto, ${eur(c.incassato)} già incassati.`),
          { cliente: nomeC(c.clienteId), chiuso: c.nome, azienda: c.azienda, funnel: c.funnel, adset: c.adset, annuncio: c.annuncio, contratto: c.valoreContratto, incassato: c.incassato }, c.clienteId);
      });
    }

    if (a.trigger === "periodo") {
      const giornoOggi = Number(g.slice(8, 10));
      if (giornoOggi < (a.giorno || 5)) return;
      data.clienti.filter((z) => z.stato === "attivo" && z.archiviato !== "si" && suo(z.id)).forEach((z) => {
        if ((data.vendite || []).some((r) => r.clienteId === z.id && r.periodo === mm)) return;
        push(`${a.id}|${z.id}|${mm}`, `Nessun dato commerciale di ${meseIt(mm)} · ${z.nome}`,
          "Il mese è iniziato da giorni e non ci sono ancora chiamate, discovery o demo registrate.",
          { cliente: z.nome, mese: mm }, z.id);
      });
    }
  });
  return out;
}

/* Il browser non può leggere la risposta di un webhook esterno (regole CORS):
   l'invio parte, ma la conferma di consegna arriva solo dal servizio ricevente. */
const inviaWebhook = async (url, payload) => {
  if (!url) return false;
  try {
    await fetch(url, { method: "POST", mode: "no-cors", headers: { "Content-Type": "text/plain;charset=UTF-8" }, body: JSON.stringify(payload) });
    return true;
  } catch (e) {
    console.error("Webhook non inviato", e);
    return false;
  }
};

/* ================================================================== */
/* Loghi                                                               */
/* ================================================================== */

const leggiLogo = (file, cb) => {
  const r = new FileReader();
  r.onload = () => {
    const im = new Image();
    im.onload = () => {
      const max = 160;
      const sc = Math.min(1, max / Math.max(im.width, im.height));
      const c = document.createElement("canvas");
      c.width = Math.round(im.width * sc); c.height = Math.round(im.height * sc);
      c.getContext("2d").drawImage(im, 0, 0, c.width, c.height);
      cb(c.toDataURL("image/png"));
    };
    im.onerror = () => cb("");
    im.src = r.result;
  };
  r.onerror = () => cb("");
  r.readAsDataURL(file);
};

const iniziali = (s) => (s || "?").split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0]).join("").toUpperCase();

const Logo = ({ src, nome, size = 28 }) =>
  src ? (
    <img className="logo" src={src} alt="" style={{ width: size, height: size }} />
  ) : (
    <span className="logo iniz" style={{ width: size, height: size, fontSize: Math.round(size * 0.36) }}>{iniziali(nome)}</span>
  );

/* ================================================================== */
/* Stile                                                               */
/* ================================================================== */

/* Icone dei riquadri: il colore serve a riconoscere il tipo di numero a colpo d'occhio,
   non a decorare — soldi che escono, volumi, costi unitari, ritorni. */

/* Un'icona per sezione: la stessa a livello agenzia e dentro il cliente,
   così la posizione nel menu cambia ma il segno resta riconoscibile. */
/* Icona disegnata a mano: Lucide non ha un semaforo, e tre pallini
   nell'ordine giusto dicono più di qualunque simbolo generico. */
const IcoSemaforo = ({ size = 16, strokeWidth = 1.3 }) => (
  <svg width={size} height={size} viewBox="0 0 16 16" fill="none" aria-hidden="true">
    <rect x="4.5" y="1.5" width="7" height="13" rx="2.2" stroke="currentColor" strokeWidth={strokeWidth} />
    <circle cx="8" cy="4.6" r="1.45" fill="#D8455F" />
    <circle cx="8" cy="8" r="1.45" fill="#C99A3A" />
    <circle cx="8" cy="11.4" r="1.45" fill="#3FA45C" />
  </svg>
);

const ICONE_NAV = {
  semaforo: IcoSemaforo, pixel: Code2, calendario: CalendarRange, cruscotto: LayoutDashboard, clienti: Building2, sottoclienti: Users, advertising: Megaphone, campagne: Megaphone,
  mesi: CalendarRange, avvisi: AlertTriangle, grafici: PieChart, vendite: PhoneCall,
  chiusure: Handshake, pianificatore: Calculator, funnel: Filter, diario: NotebookPen,
  file: FolderOpen, economia: Euro, ai: Sparkles, task: ListChecks, automazioni: Zap,
  importa: Upload, backup: DownloadCloud, guida: HelpCircle, impostazioni: Settings, esci: LogOut,
  contatti: Contact, pipeline: KanbanSquare, messaggio: PenLine, report: FileText,
  scheda: Receipt, winner: Trophy, academy: GraduationCap,
};

const IcoNav = ({ id }) => {
  const I = ICONE_NAV[id] || BarChart3;
  return <I size={16} strokeWidth={1.9} aria-hidden="true" />;
};

const ICONE_KPI = [
  { re: /avvis|critic|attenzione|rischio|concentrazion/i, I: AlertTriangle, c: "#D8455F" },
  { re: /ritorno sulla spesa|roas|resa|rendimento/i, I: TrendingUp, c: "#3FA45C" },
  { re: /costo per conversione|costo per|cpl|cpa|cac|costo chius/i, I: Target, c: "#C99A3A" },
  { re: /tempo al primo|ore |giorni|scadenz|task/i, I: Clock, c: "#C99A3A" },
  { re: /arrivano in fondo|tasso|percentu|quota|show/i, I: Percent, c: "#4C8DF6" },
  { re: /spes|budget|investit|costo orario/i, I: Wallet, c: "#D8455F" },
  { re: /incass|ricav|contratt|fattur|ricorrente|margine|profit|entrano/i, I: Euro, c: "#3FA45C" },
  { re: /roas|ritorno|resa|rendimento/i, I: TrendingUp, c: "#3FA45C" },
  { re: /chius|vendit|conversion|risultati|vinte/i, I: CheckCircle2, c: "#3FA45C" },
  { re: /cpl|cpa|cac|costo per|costo chius/i, I: Target, c: "#C99A3A" },
  { re: /lead|contatt|clienti|pubblico/i, I: Users, c: "#4C8DF6" },
  { re: /ctr|clic|click|impression/i, I: MousePointerClick, c: "#8B5CF6" },
  { re: /task|scadenz|ore|tempo|giorni/i, I: Clock, c: "#C99A3A" },
  { re: /capacit|slot|frequenz|ritmo|erogazion/i, I: Gauge, c: "#8B5CF6" },
  { re: /pipeline|trattativ|portafoglio/i, I: Briefcase, c: "#4C8DF6" },
  { re: /campagn|funnel|creativ|annunci/i, I: Layers, c: "#8B5CF6" },
  { re: /tasso|percentu|quota|show/i, I: Percent, c: "#4C8DF6" },
];

const IcoKpi = ({ t = "" }) => {
  const m = ICONE_KPI.find((x) => x.re.test(t)) || { I: BarChart3, c: "#8A8A96" };
  const I = m.I;
  return (
    <span className="ic" style={{ color: m.c, background: `${m.c}1F`, boxShadow: `inset 0 0 0 1px ${m.c}33` }}>
      <I size={15} strokeWidth={2} />
    </span>
  );
};

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');
.crm{--nero:#08080A;--ink:#F2F2F5;--ink2:#A0A0AC;--ink3:#6C6C79;--paper:#0B0B0E;--surf:#131317;--surf2:#191920;
--line:#26262E;--acc:#8E2039;--acc-t:#D8455F;--acc-d:#701829;--acc-s:rgba(142,32,57,.16);--acc-b:rgba(142,32,57,.5);
--ok:#3FA45C;--warn:#C99A3A;
--ui:"Inter","SF Pro Text",-apple-system,"Segoe UI Variable Text","Segoe UI",Roboto,ui-sans-serif,system-ui,sans-serif;
--mono:"JetBrains Mono","SF Mono",ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;
background:var(--paper);color:var(--ink);font-family:var(--ui);min-height:100%;font-size:14px;line-height:1.5;
letter-spacing:-.006em;font-feature-settings:"cv05" 1,"cv11" 1,"ss01" 1;
-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;text-rendering:optimizeLegibility}
.crm *{box-sizing:border-box}
/* Cifre della stessa larghezza ovunque: le colonne di numeri restano in colonna
   anche quando i valori cambiano, e le tabelle non "ballano" mentre si aggiornano. */
.crm{font-feature-settings:"cv05" 1,"ss01" 1;font-variant-numeric:tabular-nums}
.crm .mono,.crm .tbl .r,.crm .kpi .v,.crm .eyebrow,.crm .pill,.crm input.num{font-variant-numeric:tabular-nums lining-nums}
.crm h1,.crm h2,.crm .kpi .v{font-feature-settings:"cv05" 1,"ss01" 1,"case" 1}
.crm button{font:inherit;color:inherit;cursor:pointer;border:0;background:none}
.crm input,.crm select,.crm textarea{font:inherit;color:inherit;width:100%;background:#0F0F13;
border:1px solid var(--line);border-radius:6px;padding:7px 9px}
.crm input[type=file]{padding:5px;font-size:12px}
.crm input:focus,.crm select:focus,.crm textarea:focus{outline:1px solid var(--acc-t);outline-offset:0;border-color:var(--acc-t)}
.crm ::placeholder{color:#55555F}
/* Pannello filtri: si apre sotto l'intestazione, in griglia, con le etichette */
.pannello-filtri{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;
padding:15px 16px;border-bottom:1px solid var(--line);background:#101014}
.pannello-filtri .f{gap:5px}
.pannello-filtri label{font-size:10.5px}
.pannello-filtri select{font-size:12.8px}
.pannello-filtri .hint{font-size:11px}
/* Scelta multipla: tre colonne allineate a sinistra, così i nomi e gli stati
   restano incolonnati anche quando i testi hanno lunghezze diverse. */
.sceltatesta{display:flex;align-items:center;gap:6px;padding:0 0 8px;border-bottom:1px solid var(--line)}
.sceltalista{display:flex;flex-direction:column;max-height:260px;overflow:auto;margin:0 -2px}
.sceltariga{display:grid;grid-template-columns:18px minmax(0,1fr) 92px minmax(0,150px);
align-items:center;gap:10px;padding:7px 2px;border-bottom:1px solid rgba(255,255,255,.05);cursor:pointer;
text-align:left}
.sceltariga:last-child{border-bottom:0}
.sceltarigaintest{display:grid;grid-template-columns:18px minmax(0,1fr) 92px minmax(0,150px);
align-items:center;gap:10px;padding:0 2px 6px;font-family:var(--mono);font-size:9.5px;
letter-spacing:.14em;text-transform:uppercase;color:var(--ink3)}
.sceltarigaintest.larga{grid-template-columns:18px minmax(0,1fr) 108px 86px 78px minmax(0,150px)}
.sceltariga.larga{grid-template-columns:18px minmax(0,1fr) 108px 86px 78px minmax(0,150px)}
.sceltariga:hover{background:rgba(255,255,255,.035)}
.sceltariga input{width:16px;height:16px;margin:0;padding:0}
.sceltariga .nm{font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sceltariga[data-on="1"] .nm{color:#fff;font-weight:550}
.sceltariga .st,.sceltariga .alt{font-size:11.5px;color:var(--ink3);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sceltariga .alt{color:var(--acc-t)}
.sceltariga .mono{font-family:var(--mono);font-variant-numeric:tabular-nums;text-align:right}
.sceltarigaintest.larga span:nth-child(4),.sceltarigaintest.larga span:nth-child(5){text-align:right}
.sceltariga .st .tag{font-size:9.5px;padding:2px 6px}
.drop{border:1px dashed var(--line);border-radius:10px;padding:14px;display:grid;gap:8px;
background:#0F0F13;transition:border-color .12s ease,background .12s ease}
.drop[data-sopra="1"]{border-color:var(--acc-t);background:var(--acc-s)}
.drop input[type="file"]{border:0;background:transparent;padding:0}
.drop .sub{font-size:12px}
.crm ::-webkit-scrollbar{width:10px;height:10px}
.crm ::-webkit-scrollbar-thumb{background:#22222A;border-radius:6px;border:2px solid var(--paper)}
.crm ::-webkit-scrollbar-track{background:transparent}
.crm button:focus-visible{outline:2px solid var(--acc);outline-offset:2px;border-radius:6px}

.eyebrow{font-family:var(--mono);font-size:9.5px;font-weight:500;letter-spacing:.16em;text-transform:uppercase;color:var(--ink3)}
.mono{font-family:var(--mono);font-variant-numeric:tabular-nums slashed-zero;letter-spacing:-.02em}
.neg{color:var(--acc-t)}
.pos{font-weight:600;color:var(--ok)}
.sub{color:var(--ink2);font-size:12px}
.clic{cursor:pointer}
.wrap{overflow-x:auto}
.riga{display:flex;align-items:center;gap:10px}

.logo{object-fit:contain;background:#EDEDF0;flex:none;display:inline-block;border-radius:50%}
.logo.iniz{display:inline-flex;align-items:center;justify-content:center;background:var(--acc);color:#fff;
font-family:var(--mono);letter-spacing:.02em}

.shell{display:flex;min-height:100vh;align-items:stretch}
.rail{width:214px;flex:none;background:var(--nero);border-right:1px solid var(--line);color:#9A9AA6;padding:16px 12px;display:flex;flex-direction:column;gap:2px;
position:sticky;top:0;height:100vh;overflow-y:auto;overflow-x:hidden}
.brand{display:flex;align-items:center;gap:10px;padding:2px 8px 14px}
.brand b{font-size:15px;color:#fff;font-weight:600;display:block;line-height:1.2;letter-spacing:-.02em}
.brand span{font-family:var(--mono);font-size:9.5px;letter-spacing:.14em;text-transform:uppercase;color:#5F5F6B}
.switch{margin-bottom:12px;padding:0 2px}
.switch select{background:#15151A;border-color:#2A2A33;color:#fff;font-size:12.5px;padding:8px 9px}
.switch .eyebrow{display:block;margin:0 0 5px 8px;color:#5F5F6B}
.navb{display:flex;align-items:center;gap:9px;padding:8px 10px;border-radius:7px;font-size:13.5px;font-weight:450;letter-spacing:-.006em;width:100%;text-align:left}
.navb:hover{background:rgba(255,255,255,.05)}
.navb[data-on="1"]{background:var(--acc);color:#fff;box-shadow:0 0 0 1px rgba(216,69,95,.35)}
.navb i{font-style:normal;width:18px;height:18px;flex:none;display:inline-flex;align-items:center;justify-content:center;
color:var(--ink3);transition:color .12s ease}
.navb:hover i{color:var(--ink2)}
.navb[data-on="1"] i{color:#fff}
.navb[data-badge="1"] i{color:var(--acc-t)}
.navb[data-on="1"][data-badge="1"] i{color:#fff}
.navb u{margin-left:auto;text-decoration:none;font-family:var(--mono);font-size:11px;opacity:.55}
.navb u[data-al="1"]{background:var(--acc);color:#fff;opacity:1;border-radius:9px;padding:1px 6px}
.navb[data-on="1"] u[data-al="1"]{background:#fff;color:var(--acc)}
.railfine{margin-top:auto;display:flex;flex-direction:column;gap:2px;width:100%}
.rail[data-compatto="1"] .railfine{align-items:center}
/* Esci resta sempre visibile in fondo al menu, anche mentre le voci scorrono */
.esci-fisso{position:sticky;bottom:-16px;margin:6px -12px -16px;padding:6px 12px 16px;background:var(--nero);border-top:1px solid var(--line);width:calc(100% + 24px)}
.rail[data-compatto="1"] .esci-fisso{bottom:-14px;margin:6px -8px -14px;padding:6px 8px 14px;width:calc(100% + 16px);display:flex;justify-content:center}
.railfoot{padding:10px;font-size:11px;color:#5F5F6B;display:flex;flex-direction:column;gap:8px}

.tgl{margin-left:auto;color:#8A8A96;font-size:15px;line-height:1;padding:5px 6px;border-radius:6px;flex:none}
.tgl:hover{background:rgba(255,255,255,.07);color:#fff}
.tgl.solo{margin:0 0 10px;align-self:center}
.accbtn{width:44px;height:38px;border-radius:8px;background:#15151A;border:1px solid #2A2A33;
display:flex;align-items:center;justify-content:center;color:#fff;font-size:14px}
.accbtn:hover{border-color:var(--acc)}

.rail[data-compatto="1"]{width:64px;padding:14px 8px;align-items:center}
.rail[data-compatto="1"] .brand{padding:2px 0 10px;justify-content:center;gap:0}
.rail[data-compatto="1"] .brand .testo{display:none}
.rail[data-compatto="1"] .switch{padding:0;margin-bottom:10px;display:flex;justify-content:center}
.rail[data-compatto="1"] .navb{width:44px;justify-content:center;padding:10px 0;position:relative}
.rail[data-compatto="1"] .navb .etichetta,.rail[data-compatto="1"] .navb u{display:none}
.rail[data-compatto="1"] .navb i{width:20px;height:20px}
.rail[data-compatto="1"] .navb[data-badge="1"]::after{content:"";position:absolute;top:7px;right:7px;
width:7px;height:7px;border-radius:50%;background:var(--acc);box-shadow:0 0 0 2px var(--nero)}
.rail[data-compatto="1"] .navb[data-on="1"][data-badge="1"]::after{background:#fff;box-shadow:0 0 0 2px var(--acc)}
.rail[data-compatto="1"] .railfoot{display:none}

.main{flex:1;min-width:0;display:flex;flex-direction:column}
.top{display:flex;align-items:center;gap:12px;flex-wrap:wrap;padding:15px 22px;
border-bottom:1px solid var(--line);background:rgba(11,11,14,.82);backdrop-filter:blur(10px);
position:sticky;top:0;z-index:20}
/* Il titolo cede spazio, i comandi a destra no: sono sempre raggiungibili */
.top>div:first-of-type{min-width:0;flex:1 1 220px}
.top .lingua{flex:none}
.top h1{margin:0;font-size:17px;font-weight:600;letter-spacing:-.022em}
.top p{margin:2px 0 0;color:var(--ink2);font-size:12.5px}
.body{padding:20px 22px 56px;display:flex;flex-direction:column;gap:18px}

/* .crm davanti: senza, la regola generica ".crm button" toglieva il fondo #8E2039 ai pulsanti principali */
.crm .btn{background:var(--acc);color:#fff;padding:8px 13px;border-radius:8px;font-size:12.8px;font-weight:550;letter-spacing:-.008em;white-space:nowrap}
.crm .btn:hover{background:var(--acc-d)}
.crm .btn.ghost{background:var(--surf);color:var(--ink);border:1px solid var(--line)}
.crm .btn.ghost:hover{border-color:var(--acc-b);background:var(--surf2)}
.crm .btn.ghost[data-on="1"]{background:var(--acc);color:#fff;border-color:var(--acc)}
.crm .btn.danger{background:none;color:var(--acc);border:1px solid var(--acc-b)}
.crm .btn.danger:hover{background:var(--acc-s)}
.crm .btn.mini{padding:5px 9px;font-size:12px}
.crm .btn:disabled{opacity:.55;cursor:default}
.meta-riga{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.meta-riga select{width:auto;min-width:0;max-width:340px;flex:1 1 220px;font-size:12.5px;padding:6px 8px}
.meta-chip{display:inline-flex;align-items:center;gap:6px;background:var(--surf2);border:1px solid var(--line);border-radius:7px;padding:3px 4px 3px 8px;font-size:12.5px;margin:2px 6px 2px 0}
.meta-chip button{background:none;color:var(--ink3);padding:0 5px;font-size:14px;line-height:1}
.meta-chip button:hover{color:var(--acc-t)}
.meta-regola{display:flex;gap:8px 14px;align-items:center;flex-wrap:wrap;font-size:12.5px}
.meta-regola select{width:auto;font-size:12.5px;padding:5px 8px}
.meta-regola label{display:inline-flex;gap:6px;align-items:center;cursor:pointer}
.meta-regola input[type=checkbox],.meta-camp input[type=checkbox]{width:16px;height:16px;flex:none;accent-color:var(--acc)}
.meta-camp{max-height:330px;overflow:auto;border:1px solid var(--line);border-radius:8px;margin:8px 0}
.meta-camp label{display:flex;gap:10px;align-items:center;padding:7px 10px;border-bottom:1px solid var(--line);cursor:pointer;font-size:12.5px}
.meta-camp label:last-child{border-bottom:0}
.meta-camp .nome{flex:1;min-width:0}
.crm .adv-freccia{display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;margin-right:6px;border-radius:5px;color:var(--ink2);border:1px solid var(--line);font-size:11px;line-height:1;vertical-align:middle;background:var(--surf2)}
.crm .adv-freccia:hover{color:var(--ink);border-color:var(--acc-t)}
.adv-gruppo td{background:var(--surf2)}
.crm .acc-bloccato{color:var(--acc-t);font-size:11.5px;margin-top:3px;line-height:1.35}
.meta-msg{padding:10px 18px;font-size:13px;border-bottom:1px solid var(--line)}
.meta-msg[data-tipo="errore"]{color:var(--acc-t)}
.meta-msg[data-tipo="ok"]{color:var(--ok)}

.card{background:var(--surf);border:1px solid var(--line);border-radius:12px}
.card>header{padding:14px 16px;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.card>header h2{margin:0;font-size:13px;font-weight:600;letter-spacing:-.012em}
.card>header .sp{margin-left:auto}

.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(178px,1fr));gap:12px}
.kpi{background:var(--surf);border:1px solid var(--line);border-radius:12px;padding:15px 16px;position:relative;overflow:hidden;container-type:inline-size}
.kpi .ic{width:30px;height:30px;border-radius:9px;display:inline-flex;align-items:center;justify-content:center;
margin-bottom:11px}
.kpi .ic svg{display:block}
/* Con i decimali ovunque gli importi si allungano: il numero si stringe con il riquadro, mai oltre 25px */
.kpi .v{font-family:var(--ui);font-size:clamp(15px,11.5cqi,25px);white-space:nowrap;font-weight:600;font-variant-numeric:tabular-nums lining-nums;
letter-spacing:-.035em;line-height:1.15;margin-top:6px;color:#fff}
.kpi .s{font-size:11.5px;color:var(--ink2);margin-top:4px}

.tbl{width:100%;border-collapse:collapse}
.tbl th{text-align:left;font-family:var(--mono);font-size:9.5px;font-weight:500;letter-spacing:.14em;text-transform:uppercase;
color:var(--ink3);font-weight:500;padding:9px 14px;border-bottom:1px solid var(--line);white-space:nowrap}
.tbl td{padding:11px 14px;border-bottom:1px solid var(--line);vertical-align:middle}
.tbl tr:last-child td{border-bottom:0}
.tbl tbody tr:hover{background:rgba(255,255,255,.035)}
.tbl .r{text-align:right;font-family:var(--mono);font-size:13px;font-variant-numeric:tabular-nums slashed-zero;
letter-spacing:-.02em;white-space:nowrap}
.tbl b{font-weight:600}
.tbl tfoot td{border-top:1px solid var(--acc-b);padding:13px 14px;background:rgba(142,32,57,.09)}
.tbl tr.totale td{font-size:13.4px}
.totali{display:flex;flex-wrap:wrap;gap:0;border-bottom:1px solid var(--line);background:var(--surf2)}
.totali .tot{padding:10px 16px;border-right:1px solid var(--line);min-width:104px}
.totali .tot:last-child{border-right:0}
.totali .tot b{display:block;font-family:var(--mono);font-variant-numeric:tabular-nums;font-size:15px;margin-top:3px;letter-spacing:-.02em}

.tag{display:inline-block;font-family:var(--mono);font-size:10px;letter-spacing:.08em;text-transform:uppercase;
padding:3px 7px;border-radius:5px;background:#1E1E26;color:var(--ink2);white-space:nowrap}
.tag[data-k="attivo"],.tag[data-k="attiva"],.tag[data-k="cliente"],.tag[data-k="vinto"],.tag[data-k="fatto"]{background:rgba(63,164,92,.16);color:var(--ok);box-shadow:inset 0 0 0 1px rgba(63,164,92,.3)}
.tag[data-k="test"],.tag[data-k="onboarding"],.tag[data-k="opportunita"],.tag[data-k="qualificato"],.tag[data-k="alta"]{background:var(--acc-s);color:var(--acc)}
.tag[data-k="late"],.tag[data-k="urgente"],.tag[data-k="errore"]{background:var(--acc);color:#fff}
.tag[data-k="test"],.tag[data-k="pausa"],.tag[data-k="onboarding"]{background:rgba(201,154,58,.15);color:var(--warn);box-shadow:inset 0 0 0 1px rgba(201,154,58,.28)}

.port{display:flex;flex-direction:column}
.pr{display:flex;align-items:center;gap:14px;padding:13px 16px;border-bottom:1px solid var(--line)}
.pr:last-child{border-bottom:0}
.pr .who{width:200px;flex:none;display:flex;align-items:center;gap:10px}
.pr .who b{display:block;font-weight:600;font-size:13.5px}
.track{flex:1;height:26px;background:#1E1E26;border-radius:4px;position:relative;overflow:hidden;min-width:110px}
.fill{position:absolute;left:0;top:0;bottom:0;background:var(--nero)}
.fill[data-sotto="1"]{background:var(--acc)}
.mark{position:absolute;top:0;bottom:0;width:2px;background:var(--ink);opacity:.55}
.pr .val{width:160px;flex:none;text-align:right;font-family:var(--mono);font-variant-numeric:tabular-nums;font-size:13px}

.al{border:1px solid var(--line);border-left:4px solid var(--ink3);border-radius:9px;padding:14px 16px;background:var(--surf)}
.al[data-g="critica"]{border-left-color:var(--acc)}
.al[data-g="attenzione"]{border-left-color:var(--warn)}
.al[data-g="nota"]{border-left-color:var(--line);background:var(--surf2)}
.al[data-g="nota"] h3{font-weight:600}
.al h3{margin:0;font-size:14.5px;font-weight:600;letter-spacing:-.018em}
.al .meta{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:7px}
.al .misura{font-family:var(--mono);font-size:11px;color:var(--ink2);margin-left:auto;white-space:nowrap}
.al .misura b{color:var(--acc-t);font-weight:600}
.al dl{margin:11px 0 0;display:grid;grid-template-columns:112px 1fr;gap:7px 14px}
.al dt{font-family:var(--mono);font-size:10px;letter-spacing:.11em;text-transform:uppercase;color:var(--ink3);padding-top:2px}
.al dd{margin:0;font-size:13.3px}
.alerts{display:flex;flex-direction:column;gap:12px}

.scheda{background:linear-gradient(160deg,#17171E,#101015);border:1px solid var(--line);color:var(--ink);border-radius:10px;padding:20px 22px;display:flex;gap:18px;flex-wrap:wrap;align-items:flex-start}
.scheda h2{margin:6px 0 0;font-size:22px;font-weight:650;letter-spacing:-.02em}
.scheda .met{margin-left:auto;display:flex;gap:26px}
.scheda .met div{text-align:right}
.scheda .met .v{font-family:var(--mono);font-size:19px;font-variant-numeric:tabular-nums;margin-top:5px}
.scheda .eyebrow{color:var(--ink3)}

.board{display:grid;grid-template-columns:repeat(4,minmax(180px,1fr));gap:12px}
.col{background:var(--surf);border:1px solid var(--line);border-radius:10px;display:flex;flex-direction:column;min-height:110px}
.col>h3{margin:0;padding:11px 13px;font-size:12px;font-weight:600;display:flex;gap:7px;border-bottom:1px solid var(--line)}
.col>h3 span{margin-left:auto;font-family:var(--mono);font-size:11px;color:var(--ink2);font-weight:400}
.col .list{padding:9px;display:flex;flex-direction:column;gap:8px}
.deal{border:1px solid var(--line);border-left:3px solid var(--acc);border-radius:8px;padding:10px;background:var(--surf2)}
.deal b{display:block;font-size:13px;font-weight:600}
.deal .row{display:flex;align-items:center;gap:6px;margin-top:8px}
.deal .val{font-family:var(--mono);font-size:13px;font-variant-numeric:tabular-nums}
.move{border:1px solid var(--line);border-radius:6px;width:24px;height:24px;font-size:12px;color:var(--ink2);background:var(--surf)}
.move:hover:not(:disabled){background:var(--acc);color:#fff;border-color:var(--acc)}
.move:disabled{opacity:.3;cursor:not-allowed}

.filtri{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.filtri select{width:auto;min-width:118px;font-size:12.5px;padding:6px 8px}
.tk{border:1px solid var(--line);border-left:3px solid var(--ink3);border-radius:8px;padding:10px;background:var(--surf2)}
.tk[data-p="urgente"]{border-left-color:var(--acc)}
.tk[data-p="alta"]{border-left-color:var(--acc-t)}
.tk[data-p="media"]{border-left-color:var(--nero)}
.tk[data-scaduta="1"]{background:var(--acc-s)}
.tk b{display:block;font-size:13px;font-weight:600}
.tk .meta{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-top:7px}
.tk .row{display:flex;align-items:center;gap:6px;margin-top:8px}
.pill{font-family:var(--mono);font-size:10px;color:var(--ink2);background:#1E1E26;border-radius:5px;padding:2px 6px;white-space:nowrap}
.pill[data-late="1"]{background:var(--acc);color:#fff}
.gruppo{padding:12px 16px 4px;font-family:var(--mono);font-size:10px;letter-spacing:.14em;text-transform:uppercase;color:var(--ink3)}
.tkrow{display:flex;align-items:flex-start;gap:11px;padding:10px 16px;border-bottom:1px solid var(--line)}
.tkrow:last-child{border-bottom:0}
.tkrow input[type=checkbox]{width:16px;height:16px;margin-top:2px;flex:none;accent-color:var(--acc)}
.tkrow .t{flex:1;min-width:0}
.tkrow[data-done="1"] .t b{text-decoration:line-through;color:var(--ink3);font-weight:500}
.tkrow .meta{display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.barra{height:4px;background:#1E1E26;border-radius:2px;overflow:hidden;margin-top:7px}
.barra i{display:block;height:100%;background:var(--nero)}

.acts{display:flex;flex-direction:column}
.empty{padding:30px 16px;text-align:center;color:var(--ink2);font-size:13px}
.empty b{display:block;color:var(--ink);font-size:14px;margin-bottom:4px}

.scrim{position:fixed;inset:0;background:rgba(4,4,6,.72);backdrop-filter:blur(3px);display:flex;align-items:flex-start;justify-content:center;
padding:26px 16px;overflow:auto;z-index:40}
.modal{background:var(--surf);border-radius:12px;width:100%;max-width:640px;box-shadow:0 24px 60px rgba(0,0,0,.6);border:1px solid var(--line)}
.modal>header{padding:16px 18px;border-bottom:1px solid var(--line)}
.modal>header h2{margin:0;font-size:15px;font-weight:650}
.modal>header p{margin:3px 0 0;font-size:12px;color:var(--ink2)}
.fields{padding:16px 18px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:13px}
.f{display:flex;flex-direction:column;gap:5px}
.f label{font-family:var(--mono);font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--ink3)}
.f input.num{font-family:var(--mono);font-variant-numeric:tabular-nums;text-align:right}
.f .hint{font-size:11px;color:var(--ink3)}
.f.w2{grid-column:span 2}
.f.wide{grid-column:1/-1}
.grp{grid-column:1/-1;font-family:var(--mono);font-size:10px;letter-spacing:.14em;text-transform:uppercase;color:var(--ink);
border-bottom:1px solid var(--line);padding-bottom:6px;margin-top:6px}
.grp span{color:var(--ink3);letter-spacing:0;text-transform:none;font-family:var(--ui);font-size:11.5px;margin-left:8px}
.logoedit{display:flex;align-items:center;gap:12px}
.logoedit .cmd{flex:1;display:flex;flex-direction:column;gap:6px;min-width:0}
.ck{display:flex;flex-direction:column;gap:6px}
.ckr{display:flex;align-items:center;gap:8px}
.ckr input[type=checkbox]{width:16px;height:16px;flex:none;accent-color:var(--acc)}
.ckr button{color:var(--ink3);font-size:16px;line-height:1;padding:0 4px}
.ckr button:hover{color:var(--acc)}
.pr-riga{display:grid;grid-template-columns:1.6fr 1.1fr .7fr .7fr auto;gap:6px;align-items:center}
.pr-riga button{color:var(--ink3);font-size:16px;line-height:1;padding:0 4px}
.pr-riga button:hover{color:var(--acc)}
@media (max-width:700px){.pr-riga{grid-template-columns:1fr 1fr;gap:5px}.pr-riga button{grid-column:2;justify-self:end}}
.modal>footer{padding:13px 18px;border-top:1px solid var(--line);display:flex;gap:8px;align-items:center}
.modal>footer .sp{margin-left:auto}

.two{display:grid;grid-template-columns:1.25fr .75fr;gap:18px}
.saved{font-family:var(--mono);font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--acc)}
.funnel{width:100%;height:auto;display:block}

.ctrl{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.ctrl select{width:auto;min-width:126px;font-size:12.5px;padding:6px 8px}
.bars{display:flex;flex-direction:column;gap:9px;padding:16px}
.bar{display:flex;align-items:center;gap:12px}
.bar .et{width:190px;flex:none;font-size:12.5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bar .tr{flex:1;height:22px;background:#1E1E26;border-radius:3px;overflow:hidden;min-width:60px}
.bar .tr i{display:block;height:100%;border-radius:3px}
.bar .vv{width:118px;flex:none;text-align:right;font-family:var(--mono);font-size:12.5px;font-variant-numeric:tabular-nums}
.pie{display:flex;gap:24px;align-items:center;padding:16px;flex-wrap:wrap}
.pie svg{flex:none}
.leg{flex:1;min-width:210px;display:flex;flex-direction:column;gap:8px}
.lg{display:flex;align-items:center;gap:9px;font-size:12.5px}
.lg u{width:10px;height:10px;border-radius:3px;flex:none;text-decoration:none}
.lg .n{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.lg .v{font-family:var(--mono);font-variant-numeric:tabular-nums;font-size:12px;white-space:nowrap}
.nota{padding:0 16px 14px;font-size:12px;color:var(--ink2)}

.report{display:flex;flex-direction:column;gap:18px}

.chat{display:flex;flex-direction:column;gap:12px;padding:16px;max-height:min(58vh,560px);overflow-y:auto}
.msg{max-width:min(78ch,88%);border-radius:12px;padding:12px 14px;font-size:13.6px;line-height:1.55;white-space:pre-wrap}
.msg[data-da="io"]{align-self:flex-end;background:var(--acc);color:#fff;border-bottom-right-radius:4px}
.msg[data-da="ai"]{align-self:flex-start;background:var(--surf2);border:1px solid var(--line);border-bottom-left-radius:4px}
.msg .azioni{display:flex;gap:6px;margin-top:10px;flex-wrap:wrap}
.msg .puntini{display:inline-flex;gap:4px}
.msg .puntini i{width:6px;height:6px;border-radius:50%;background:var(--acc);opacity:.35;animation:bat 1.1s infinite}
.msg .puntini i:nth-child(2){animation-delay:.18s}
.msg .puntini i:nth-child(3){animation-delay:.36s}
@keyframes bat{0%,100%{opacity:.25}50%{opacity:1}}
.compositore{display:flex;gap:8px;align-items:flex-end;padding:12px 16px;border-top:1px solid var(--line)}
.compositore textarea{resize:vertical;min-height:44px}
.rtesta{display:flex;align-items:center;gap:16px;flex-wrap:wrap;background:var(--surf);
border:1px solid var(--line);border-radius:10px;padding:18px 20px}
.rtesta h1{margin:0;font-size:22px;font-weight:600;letter-spacing:-.028em}

@media print{
 .crm{background:#fff;color:#131316;--paper:#fff;--surf:#fff;--surf2:#F7F5F6;--line:#DAD5D6;--ink:#131316;--ink2:#565660;--ink3:#8C8C96}
 .card,.kpi{background:#fff}
 .kpi .v{color:#131316}
 .tbl tfoot td{background:#F7F5F6}
 .scheda{background:#F7F5F6;color:#131316}
 .rail,.top,.nostampa{display:none!important}
 .body{padding:0!important;gap:14px}
 .shell{display:block}
 .card,.kpi,.rtesta{border-color:#CFCACB;break-inside:avoid}
 .two{grid-template-columns:1fr 1fr}
 .kpis{grid-template-columns:repeat(4,1fr)}
 .report{gap:14px}
}

@media (max-width:900px){
 .railfoot{display:none}
 .bar .et{width:104px}
 .bar .vv{width:86px;font-size:12px}
 .pie{gap:14px}
 .kpis{grid-template-columns:1fr 1fr}
 .two,.board{grid-template-columns:1fr}
 .fields{grid-template-columns:1fr 1fr}
 .f.w2,.f.wide{grid-column:1/-1}
 .pr{flex-wrap:wrap}
 .pr .who{width:100%}
 .pr .val{width:auto;margin-left:auto}
 .al dl{grid-template-columns:1fr;gap:3px 0}
 .al dd{margin-bottom:7px}
 .body,.top{padding-left:14px;padding-right:14px}
}
@media (max-width:560px){.kpis{grid-template-columns:1fr}}
@media (prefers-reduced-motion:reduce){.crm *{transition:none!important}}
/* Calendario */
.cal-legenda{display:flex;align-items:center;gap:14px;flex-wrap:wrap;padding:11px 16px;border-bottom:1px solid var(--line)}
.cal-legenda label{display:inline-flex;align-items:center;gap:7px;font-size:12px;color:var(--ink3);cursor:pointer}
.cal-legenda label.on{color:var(--ink2)}
.cal-legenda label i{width:9px;height:9px;border-radius:3px;display:inline-block;opacity:.45}
.cal-legenda label.on i{opacity:1}
.cal-legenda input{width:14px;height:14px;flex:none}
.cal-mese{display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--line);padding:1px}
.cal-testa{background:#0F0F13;font-family:var(--mono);font-size:9.5px;letter-spacing:.14em;text-transform:uppercase;
color:var(--ink3);padding:8px 10px;display:flex;align-items:center;justify-content:space-between}
.cal-cella{background:var(--surf);min-height:112px;padding:7px 8px;display:flex;flex-direction:column;gap:3px;
transition:background .12s ease}
.cal-cella:hover{background:#121217}
.cal-cella.fuori{background:#0E0E12;opacity:.5}
.cal-cella.oggi{box-shadow:inset 0 0 0 1px var(--acc)}
.cal-cella.oggi .n span{background:var(--acc);color:#fff;border-radius:50%;width:19px;height:19px;
display:inline-flex;align-items:center;justify-content:center}
.cal-cella .n{display:flex;align-items:center;justify-content:space-between;font-family:var(--mono);font-size:11px;color:var(--ink3)}
.cal-cella .piu,.cal-testa .piu{opacity:0;background:transparent;color:var(--ink3);font-size:14px;line-height:1;padding:0 3px}
.cal-cella:hover .piu,.cal-testa:hover .piu{opacity:1}
.cal-cella .piu:hover,.cal-testa .piu:hover{color:#fff}
.cal-voce{display:block;width:100%;text-align:left;background:rgba(255,255,255,.045);border-left:2.5px solid var(--ink3);
border-radius:5px;padding:3.5px 7px;font-size:11.5px;line-height:1.3;overflow:hidden;text-overflow:ellipsis;
white-space:nowrap;transition:background .1s ease}
.cal-voce:hover{background:rgba(255,255,255,.09)}
.cal-voce em{font-style:normal;color:var(--ink3);margin-left:5px;font-size:10.5px}
.cal-altro{font-size:10.5px;color:var(--ink3);background:transparent;text-align:left;padding:1px 6px}
.cal-altro:hover{color:#fff}
.cal-sett{display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--line);padding:1px}
.cal-colonna{background:var(--surf);min-height:260px;display:flex;flex-direction:column}
.cal-colonna.oggi{box-shadow:inset 0 0 0 1px var(--acc)}
.cal-colonna .corpo{padding:7px;display:flex;flex-direction:column;gap:4px}
.cal-colonna .vuoto{font-size:11px;padding:2px 2px}
.cal-cat{display:inline-flex;align-items:center;gap:8px;background:var(--surf2);border:1px solid var(--line);
border-radius:999px;padding:6px 13px;font-size:12.5px}
.cal-cat:hover{border-color:var(--acc-b)}
.cal-cat i{width:10px;height:10px;border-radius:3px}
/* Vista giornaliera: fascia di testa e griglia oraria */
.cal-giorno{padding:0 0 10px}
.tuttoilgiorno{display:grid;grid-template-columns:76px 1fr;gap:12px;align-items:start;
padding:12px 16px;border-bottom:1px solid var(--line);background:#0F0F13}
.tuttoilgiorno .et{font-family:var(--mono);font-size:9.5px;letter-spacing:.12em;text-transform:uppercase;
color:var(--ink3);padding-top:4px}
.tuttoilgiorno .voci{display:flex;flex-direction:column;gap:4px;min-width:0}
.ore{display:flex;flex-direction:column}
.fascia{display:grid;grid-template-columns:76px 1fr;gap:12px;align-items:stretch;position:relative;
border-bottom:1px solid rgba(255,255,255,.05);min-height:46px}
.fascia:hover{background:rgba(255,255,255,.02)}
.fascia .h{font-family:var(--mono);font-size:11px;color:var(--ink3);padding:9px 0 0 16px;text-align:right;
padding-right:0}
.fascia .slot{display:flex;flex-direction:column;gap:4px;padding:7px 16px 7px 0;position:relative}
.fascia .blocco{border-left:3px solid var(--ink3);border-radius:7px;display:flex;align-items:stretch}
.fascia .blocco button{flex:1;text-align:left;background:transparent;padding:7px 11px;font-size:12.5px;line-height:1.35}
.fascia .blocco b{font-family:var(--mono);font-size:11.5px;margin-right:6px;color:#fff}
.fascia .blocco em{font-style:normal;color:var(--ink3);font-size:11.5px}
.fascia .aggiungi{align-self:flex-start;opacity:0;background:transparent;color:var(--ink3);font-size:15px;
line-height:1;padding:2px 6px}
.fascia:hover .aggiungi{opacity:1}
.fascia .aggiungi:hover{color:#fff}
.fascia .adesso{position:absolute;left:76px;right:0;height:1px;background:var(--acc);pointer-events:none}
.fascia .adesso::before{content:"";position:absolute;left:-4px;top:-3px;width:7px;height:7px;border-radius:50%;background:var(--acc)}

/* Vista annuale: dodici mesi veri */
.cal-anno{display:grid;grid-template-columns:repeat(auto-fit,minmax(196px,1fr));gap:12px;padding:16px}
.mesetto{background:var(--surf2);border:1px solid var(--line);border-radius:11px;padding:11px 12px 12px}
.mesetto header{display:flex;align-items:baseline;justify-content:space-between;padding:0 0 8px;border:0;background:none}
.mesetto header b{font-size:13px;text-transform:capitalize;cursor:pointer}
.mesetto header b:hover{color:var(--acc-t)}
.mesetto .n{font-size:11px;color:var(--ink3)}
.cal-mini{display:grid;grid-template-columns:repeat(7,1fr);gap:2px}
.cal-mini .dow{font-family:var(--mono);font-size:8.5px;color:#4A4A55;text-align:center;padding-bottom:3px}
.cal-mini .gg{position:relative;aspect-ratio:1;border-radius:5px;background:transparent;color:var(--ink3);
font-size:10.5px;display:flex;align-items:center;justify-content:center;padding:0}
.cal-mini .gg:hover{background:rgba(255,255,255,.07);color:#fff}
.cal-mini .gg.fuori{opacity:.22}
.cal-mini .gg.oggi{box-shadow:inset 0 0 0 1px var(--acc);color:#fff}
.cal-mini .punti{position:absolute;bottom:2px;left:0;right:0;display:flex;gap:1.5px;justify-content:center}
.cal-mini .punti i{width:3.5px;height:3.5px;border-radius:50%}
/* Pixel facile: il blocco di codice */
.px-out{margin:0;padding:18px 20px;background:#0B0B0F;border-top:1px solid var(--line);
border-bottom:1px solid var(--line);overflow-x:auto}
.px-out code{font-family:var(--mono);font-size:12.8px;line-height:1.7;color:#C8C8D2;white-space:pre;display:block}
/* Selettore della lingua, in alto a destra */
.lingua{display:inline-flex;align-items:center;gap:9px;flex:none;border:1px solid var(--line);
border-radius:9px;padding:5px 9px 5px 11px;background:var(--surf2)}
.lingua label{font-family:var(--mono);font-size:9.5px;letter-spacing:.13em;text-transform:uppercase;
color:var(--ink3);margin:0}
.lingua select{width:auto;min-width:104px;font-size:12.5px;padding:3px 24px 3px 6px;margin:0;
border:0;background-color:transparent}
.lingua select:focus{outline:none}
.lingua.valuta{border-color:var(--acc-b)}
.lingua.valuta label{color:var(--acc-t)}
/* Automazioni proposte dal consulente */
.proposta{margin-top:12px;border:1px solid var(--acc-b);border-radius:10px;padding:12px 13px;
background:var(--acc-s);display:grid;gap:9px}
.proposta .tag{flex:none}
.spunti{display:grid;gap:7px;padding:4px 0 10px;justify-items:start}
.spunti .btn{text-align:left;white-space:normal;line-height:1.4}
/* ---------- Academy in stile Netflix ---------- */
.nf{display:grid;gap:28px;grid-template-columns:minmax(0,1fr);min-width:0;max-width:100%}
.nf>*,.nf-riga,.nf-scorri{min-width:0}
.nf-hero{position:relative;border-radius:14px;overflow:hidden;min-height:360px;display:flex;align-items:flex-end;padding:36px 34px;color:#fff;background:#0C0C0E;isolation:isolate}
.nf-hero .sfondo{position:absolute;inset:0;z-index:-1}
.nf-hero .sfondo::after{content:"";position:absolute;inset:0;background:linear-gradient(90deg,rgba(8,8,10,.92) 0%,rgba(8,8,10,.55) 45%,rgba(8,8,10,0) 75%),linear-gradient(0deg,rgba(8,8,10,.85),rgba(8,8,10,0) 50%)}
.nf-hero .marchio{font-family:var(--mono);font-size:10.5px;letter-spacing:.24em;text-transform:uppercase;color:#E8A3B2;display:flex;align-items:center;gap:8px}
.nf-hero h1{font-size:clamp(28px,4.2vw,46px);line-height:1.05;letter-spacing:-.03em;margin:10px 0 10px;max-width:640px;font-weight:700}
.nf-hero p{max-width:520px;margin:0 0 18px;color:rgba(255,255,255,.82);font-size:14.5px;line-height:1.5}
.nf-hero .meta{display:flex;gap:10px;align-items:center;margin-bottom:16px;font-size:12.5px;color:rgba(255,255,255,.75)}
.crm .nf-play{background:#fff;color:#0C0C0E;border-radius:6px;padding:10px 20px;font-weight:650;font-size:14px;display:inline-flex;align-items:center;gap:8px}
.crm .nf-play:hover{background:rgba(255,255,255,.82)}
.crm .nf-info{background:rgba(110,110,120,.55);color:#fff;border-radius:6px;padding:10px 18px;font-weight:600;font-size:14px;display:inline-flex;align-items:center;gap:8px}
.crm .nf-info:hover{background:rgba(110,110,120,.4)}
.crm .nf-sblocca{background:var(--acc);color:#fff;border-radius:6px;padding:10px 18px;font-weight:650;font-size:14px}
.crm .nf-sblocca:hover{background:var(--acc-d)}
.nf-riga{position:relative}
.nf-riga h3{margin:0 0 10px;font-size:17px;font-weight:650;letter-spacing:-.015em;display:flex;align-items:baseline;gap:10px}
.nf-riga h3 small{font-size:12px;color:var(--ink3);font-weight:450}
.nf-scorri{display:grid;grid-auto-flow:column;grid-auto-columns:minmax(230px,calc((100% - 30px)/4));gap:10px;overflow-x:auto;padding:8px 2px 16px;scroll-snap-type:x mandatory;scrollbar-width:none}
.nf-scorri::-webkit-scrollbar{display:none}
.crm .nf-freccia{position:absolute;top:38px;bottom:16px;width:40px;z-index:3;background:linear-gradient(90deg,rgba(0,0,0,.55),rgba(0,0,0,0));color:#fff;font-size:26px;opacity:0;transition:opacity .15s}
.crm .nf-freccia.dx{right:0;background:linear-gradient(270deg,rgba(0,0,0,.55),rgba(0,0,0,0))}
.crm .nf-freccia.sx{left:0}
.nf-riga:hover .nf-freccia{opacity:1}
.nf-card{position:relative;aspect-ratio:16/9;border-radius:8px;overflow:hidden;cursor:pointer;scroll-snap-align:start;color:#fff;
transition:transform .2s ease,box-shadow .2s ease;background:#0C0C0E;outline-offset:2px}
.nf-card:hover,.nf-card:focus-visible{transform:scale(1.045);z-index:2;box-shadow:0 14px 34px rgba(0,0,0,.55)}
.nf-card .cop{position:absolute;inset:0}
.nf-card .cop::after{content:"";position:absolute;inset:0;background:linear-gradient(0deg,rgba(0,0,0,.85),rgba(0,0,0,0) 60%)}
.nf-card .grande{position:absolute;right:-6px;top:-18px;font-size:110px;font-weight:800;letter-spacing:-.06em;color:rgba(255,255,255,.08);line-height:1}
.nf-card .tit{position:absolute;left:12px;right:12px;bottom:12px;z-index:1;font-weight:650;font-size:14.5px;line-height:1.2;text-shadow:0 1px 6px rgba(0,0,0,.6)}
.nf-card .tit small{display:block;font-weight:450;font-size:11.5px;color:rgba(255,255,255,.75);margin-top:3px}
.nf-card .bollo{position:absolute;top:9px;left:9px;z-index:1;font-family:var(--mono);font-size:9.5px;letter-spacing:.1em;text-transform:uppercase;padding:3px 6px;border-radius:4px;background:rgba(0,0,0,.55);color:#fff}
.nf-card .bollo.gratis{background:#2E7D46}
.nf-card .bollo.chiuso{background:var(--acc)}
.nf-card .cop-img{position:absolute;inset:0;width:100%;height:100%;object-fit:cover}
.crm .nf-tutti{font-size:12.5px;font-weight:550;color:var(--acc-t);margin-left:4px}
.crm .nf-tutti:hover{text-decoration:underline}
.nf-griglia{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:14px}
.nf-schede{display:flex;gap:4px;flex-wrap:wrap;border-bottom:1px solid var(--line);padding-bottom:2px}
.crm .nf-schede button{padding:9px 14px;font-size:14px;font-weight:550;color:var(--ink2);border-bottom:2px solid transparent;margin-bottom:-3px;display:inline-flex;gap:7px;align-items:baseline}
.crm .nf-schede button small{font-family:var(--mono);font-size:10.5px;color:var(--ink3)}
.crm .nf-schede button:hover{color:var(--ink)}
.crm .nf-schede button[data-on="1"]{color:var(--ink);border-bottom-color:var(--acc)}
.nf-hero-cop{position:absolute;right:28px;top:50%;transform:translateY(-50%);width:min(440px,42%);aspect-ratio:16/9;object-fit:cover;border-radius:10px;z-index:-1;box-shadow:0 20px 50px rgba(0,0,0,.6)}
@media (max-width:900px){.nf-hero-cop{position:static;transform:none;width:100%;margin-bottom:16px;order:-1}.nf-hero{flex-direction:column;align-items:stretch}}
.nf-card .logo-n{position:absolute;top:8px;right:8px;z-index:1;width:26px;height:26px;filter:drop-shadow(0 2px 6px rgba(0,0,0,.5))}
.nf-sigillo{position:absolute;right:42px;top:50%;transform:translateY(-50%) rotate(-8deg);width:min(250px,30%);z-index:-1;opacity:.92;filter:drop-shadow(0 18px 40px rgba(0,0,0,.55))}
@media (max-width:700px){.nf-sigillo{width:110px;right:14px;top:22px;transform:rotate(-8deg);opacity:.5}}
.nf-card .avanz{position:absolute;left:0;right:0;bottom:0;height:3px;background:rgba(255,255,255,.25);z-index:1}
.nf-card .avanz i{display:block;height:100%;background:#E50914;background:var(--acc-t)}
.nf-scrim{position:fixed;inset:0;z-index:60;background:rgba(0,0,0,.72);overflow-y:auto;padding:36px 16px}
.nf-modal{max-width:880px;margin:0 auto;background:var(--surf);border-radius:12px;overflow:hidden;box-shadow:0 30px 80px rgba(0,0,0,.6);color:var(--ink)}
.nf-mtesta{position:relative;aspect-ratio:16/7;min-height:250px;display:flex;align-items:flex-end;padding:28px;color:#fff;isolation:isolate}
.nf-mtesta .sfondo{position:absolute;inset:0;z-index:-1}
.nf-mtesta .sfondo::after{content:"";position:absolute;inset:0;background:linear-gradient(0deg,var(--surf) 0%,rgba(0,0,0,.25) 45%,rgba(0,0,0,0) 100%)}
.nf-mtesta h2{font-size:32px;letter-spacing:-.03em;margin:0 0 14px;line-height:1.05;color:#fff;text-shadow:0 2px 10px rgba(0,0,0,.5)}
.crm .nf-chiudi{position:absolute;top:14px;right:14px;width:36px;height:36px;border-radius:50%;background:rgba(12,12,14,.8);color:#fff;font-size:18px;z-index:2}
.nf-corpo{padding:6px 28px 28px;display:grid;gap:16px}
.nf-dati{display:flex;gap:12px;flex-wrap:wrap;align-items:center;font-size:13px;color:var(--ink2)}
.nf-ep{display:grid;grid-template-columns:34px 150px minmax(0,1fr) auto;gap:16px;align-items:center;padding:14px 10px;border-radius:8px;cursor:pointer;width:100%;text-align:left;border-bottom:1px solid var(--line)}
.nf-ep:hover{background:var(--surf2)}
.nf-ep .n{font-size:22px;color:var(--ink3);text-align:center}
.nf-ep .mini{position:relative;aspect-ratio:16/9;border-radius:6px;overflow:hidden;display:flex;align-items:center;justify-content:center;color:#fff;font-size:20px}
.nf-ep b{display:block;font-size:14.5px}
.nf-ep .sub{font-size:12.5px}
.nf-player{aspect-ratio:16/9;width:100%;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:12px;color:#fff;position:relative}
@media (max-width:700px){.nf-scorri{grid-auto-columns:72%}.nf-ep{grid-template-columns:28px 96px minmax(0,1fr)}.nf-ep .dur{display:none}.nf-hero{padding:24px 18px;min-height:300px}}
/* ---------- Modalità chiara (Impostazioni → Aspetto): stesso #8E2039, fondi bianchi ---------- */
.crm[data-tema="chiaro"]{--nero:#FFFFFF;--ink:#17171C;--ink2:#55555F;--ink3:#85858F;--paper:#F5F4F5;--surf:#FFFFFF;--surf2:#F3F1F2;
--line:#E3DFE0;--acc-t:#B02E4B;--acc-s:rgba(142,32,57,.08);--acc-b:rgba(142,32,57,.35);color-scheme:light}
.crm[data-tema="chiaro"] input,.crm[data-tema="chiaro"] select,.crm[data-tema="chiaro"] textarea{background:#FFFFFF}
.crm[data-tema="chiaro"] ::placeholder{color:#A0A0AA}
.crm[data-tema="chiaro"] .pannello-filtri,.crm[data-tema="chiaro"] .tuttoilgiorno,.crm[data-tema="chiaro"] .cal-testa,.crm[data-tema="chiaro"] .px-out{background:#F8F7F8}
.crm[data-tema="chiaro"] .drop{background:#FAF9FA}
.crm[data-tema="chiaro"] .sceltariga,.crm[data-tema="chiaro"] .fascia{border-bottom-color:rgba(0,0,0,.06)}
.crm[data-tema="chiaro"] .sceltariga:hover,.crm[data-tema="chiaro"] .tbl tbody tr:hover{background:rgba(0,0,0,.028)}
.crm[data-tema="chiaro"] .fascia:hover{background:rgba(0,0,0,.02)}
.crm[data-tema="chiaro"] ::-webkit-scrollbar-thumb{background:#D6D1D3}
.crm[data-tema="chiaro"] .rail{color:#55555F}
.crm[data-tema="chiaro"] .brand span,.crm[data-tema="chiaro"] .switch .eyebrow,.crm[data-tema="chiaro"] .railfoot{color:#85858F}
.crm[data-tema="chiaro"] .switch select,.crm[data-tema="chiaro"] .accbtn{background:#FFFFFF;border-color:var(--line);color:var(--ink)}
.crm[data-tema="chiaro"] .navb:not([data-on="1"]):hover,.crm[data-tema="chiaro"] .tgl:hover,.crm[data-tema="chiaro"] .cal-mini .gg:hover{background:rgba(0,0,0,.05);color:var(--ink)}
.crm[data-tema="chiaro"] .tgl{color:#85858F}
.crm[data-tema="chiaro"] .top{background:rgba(245,244,245,.88)}
/* :where non aggiunge peso: le etichette colorate (urgente, alta, in ritardo…) restano più forti */
:where(.crm[data-tema="chiaro"]) .tag,:where(.crm[data-tema="chiaro"]) .pill,:where(.crm[data-tema="chiaro"]) .track,:where(.crm[data-tema="chiaro"]) .barra,:where(.crm[data-tema="chiaro"]) .bar .tr{background:#EEEBEC}
.crm[data-tema="chiaro"] .scheda{background:linear-gradient(160deg,#FFFFFF,#F6F4F5)}
.crm[data-tema="chiaro"] .scrim{background:rgba(23,23,28,.35)}
.crm[data-tema="chiaro"] .modal{box-shadow:0 24px 60px rgba(0,0,0,.16)}
.crm[data-tema="chiaro"] .cal-cella:hover{background:#FAF9FA}
.crm[data-tema="chiaro"] .cal-cella.fuori{background:#F4F3F4}
.crm[data-tema="chiaro"] .cal-voce{background:rgba(0,0,0,.04)}
.crm[data-tema="chiaro"] .cal-voce:hover{background:rgba(0,0,0,.075)}
.crm[data-tema="chiaro"] .cal-mini .dow{color:#A0A0AA}
.crm[data-tema="chiaro"] .px-out code{color:#3A3A44}
/* Testi bianchi su fondo scuro: nel chiaro diventano inchiostro (quelli su #8E2039 restano bianchi) */
.crm[data-tema="chiaro"] .sceltariga[data-on="1"] .nm,.crm[data-tema="chiaro"] .brand b,.crm[data-tema="chiaro"] .kpi .v,
.crm[data-tema="chiaro"] .cal-cella .piu:hover,.crm[data-tema="chiaro"] .cal-testa .piu:hover,.crm[data-tema="chiaro"] .cal-altro:hover,
.crm[data-tema="chiaro"] .fascia .blocco b,.crm[data-tema="chiaro"] .fascia .aggiungi:hover,.crm[data-tema="chiaro"] .cal-mini .gg.oggi{color:var(--ink)}
.crm[data-tema="chiaro"] .card,.crm[data-tema="chiaro"] .kpi{box-shadow:0 1px 2px rgba(0,0,0,.04)}
`;

/* ================================================================== */
/* Componenti                                                          */
/* ================================================================== */

function Funnel({ steps }) {
  const W = 1000, H = 150, cy = 66, gap = 46;
  const seg = (W - gap * (steps.length - 1)) / steps.length;
  const max = Math.max(...steps.map((s) => s.n), 1);
  const h = (n) => Math.max(10, (n / max) * 96);
  const tint = ["#131316", "#4A1B26", "#6E1E31", "#8E2039"];
  return (
    <svg className="funnel" viewBox={`0 0 ${W} ${H}`} role="img" aria-label="Imbuto dai lead ai clienti">
      {steps.map((s, i) => {
        const x = i * (seg + gap), hl = h(s.n), hr = h(steps[i + 1] ? steps[i + 1].n : s.n * 0.82);
        const conv = i < steps.length - 1 && s.n > 0 ? (steps[i + 1].n / s.n) * 100 : null;
        return (
          <g key={s.k}>
            <polygon points={`${x},${cy - hl / 2} ${x + seg},${cy - hr / 2} ${x + seg},${cy + hr / 2} ${x},${cy + hl / 2}`} fill={tint[i]} />
            <text x={x + 12} y={cy + 6} fill="#fff" fontFamily="var(--mono)" fontSize="19">{int(s.n)}</text>
            <text x={x} y={H - 24} fill="var(--ink)" fontSize="14" fontWeight="600">{s.k}</text>
            <text x={x} y={H - 6} fill="var(--ink3)" fontFamily="var(--mono)" fontSize="11" letterSpacing="1.4">{s.sub}</text>
            {conv !== null && <text x={x + seg + gap / 2} y={cy + 4} textAnchor="middle" fill="var(--ink2)" fontFamily="var(--mono)" fontSize="12">{fmt(conv, 1)}%</text>}
          </g>
        );
      })}
    </svg>
  );
}

function Avviso({ a, dove }) {
  return (
    <article className="al" data-g={a.gravita}>
      <div className="meta">
        <span className="tag" data-k={a.gravita === "critica" ? "late" : ""}>{a.gravita}</span>
        {dove && <span className="sub">{dove}</span>}
        <span className="misura">{a.metrica} <b>{a.valore}</b> · obiettivo {a.soglia}</span>
      </div>
      <h3>{a.titolo}</h3>
      <dl>
        <dt>Cosa succede</dt><dd>{a.cosa}</dd>
        <dt>Causa probabile</dt><dd>{a.causa}</dd>
        <dt>Cosa fare</dt><dd>{a.azione}</dd>
      </dl>
    </article>
  );
}

const SERIE = ["#0C0C0E", "#8E2039", "#3E1720", "#B4566B", "#5C1B2C", "#C99BA7", "#6E6A6C", "#E0C3CB"];

function Barre({ dati, fmtv }) {
  if (!dati.length) return <div className="empty"><b>Nessun dato</b>Non ci sono numeri da rappresentare con questa scelta.</div>;
  const max = Math.max(...dati.map((d) => Math.abs(d.v)), 1);
  return (
    <div className="bars">
      {dati.map((d, i) => (
        <div className="bar" key={d.k + i}>
          <span className="et" title={d.k}>{d.k}</span>
          <span className="tr"><i style={{ width: `${(Math.abs(d.v) / max) * 100}%`, background: SERIE[i % SERIE.length] }} /></span>
          <span className="vv">{fmtv(d.v)}</span>
        </div>
      ))}
    </div>
  );
}

function Torta({ dati, fmtv }) {
  const validi = dati.filter((d) => d.v > 0);
  const tot = validi.reduce((s, d) => s + d.v, 0);
  if (!tot) return <div className="empty"><b>Nessun dato</b>Servono valori positivi per costruire la torta.</div>;
  const C = 110, R = 92, r = 54;
  let ang = -Math.PI / 2;
  const fette = validi.map((d, i) => {
    const a0 = ang, a1 = ang + (d.v / tot) * 2 * Math.PI;
    ang = a1;
    const p = (rad, a) => `${(C + rad * Math.cos(a)).toFixed(2)},${(C + rad * Math.sin(a)).toFixed(2)}`;
    const grande = a1 - a0 > Math.PI ? 1 : 0;
    const dd = `M${p(R, a0)} A${R},${R} 0 ${grande} 1 ${p(R, a1)} L${p(r, a1)} A${r},${r} 0 ${grande} 0 ${p(r, a0)} Z`;
    return { d, dd, col: SERIE[i % SERIE.length], quota: d.v / tot };
  });
  return (
    <div className="pie">
      <svg width="220" height="220" viewBox="0 0 220 220" role="img" aria-label="Grafico a torta">
        {validi.length === 1 ? (
          <>
            <circle cx={C} cy={C} r={(R + r) / 2} fill="none" stroke={SERIE[0]} strokeWidth={R - r} />
          </>
        ) : fette.map((f, i) => <path key={i} d={f.dd} fill={f.col} />)}
        <text x={C} y={C - 4} textAnchor="middle" fontFamily="var(--mono)" fontSize="10" letterSpacing="1.4" fill="var(--ink3)">TOTALE</text>
        <text x={C} y={C + 16} textAnchor="middle" fontFamily="var(--mono)" fontSize="16" fill="var(--ink)">{fmtv(tot)}</text>
      </svg>
      <div className="leg">
        {fette.map((f, i) => (
          <div className="lg" key={i}>
            <u style={{ background: f.col }} />
            <span className="n" title={f.d.k}>{f.d.k}</span>
            <span className="v">{fmtv(f.d.v)}</span>
            <span className="v sub">{fmt(f.quota * 100, 1)}%</span>
          </div>
        ))}
      </div>
    </div>
  );
}

/* Andamento nel tempo: due serie sovrapposte, area piena per i contratti e
   linea sottile per la spesa. Niente griglia pesante: si legge la forma. */
function Andamento({ righe, alto = 210 }) {
  if (righe.length < 2) return <div className="empty">Servono almeno due mesi di dati per disegnare l'andamento.</div>;
  const L = 46, R = 14, T = 14, B = 26, W = 720, H = alto;
  const max = Math.max(...righe.map((r) => Math.max(r.contratto, r.spesa)), 1);
  const px = (i) => L + (i * (W - L - R)) / Math.max(1, righe.length - 1);
  const py = (v) => T + (H - T - B) * (1 - v / max);
  const linea = (k) => righe.map((r, i) => `${i ? "L" : "M"}${px(i).toFixed(1)} ${py(r[k]).toFixed(1)}`).join(" ");
  const area = `${linea("contratto")} L${px(righe.length - 1).toFixed(1)} ${py(0)} L${px(0).toFixed(1)} ${py(0)} Z`;
  const tacche = [max, max / 2, 0];
  return (
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" height={H} role="img" aria-label="Andamento di contratti e spesa">
      <defs>
        <linearGradient id="gAnd" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor="#8E2039" stopOpacity="0.45" />
          <stop offset="100%" stopColor="#8E2039" stopOpacity="0.02" />
        </linearGradient>
      </defs>
      {tacche.map((t, i) => (
        <g key={i}>
          <line x1={L} y1={py(t)} x2={W - R} y2={py(t)} stroke="var(--line)" strokeWidth="1" />
          <text x={L - 8} y={py(t) + 4} textAnchor="end" fontSize="10" fill="#6C6C79" fontFamily="JetBrains Mono, ui-monospace, monospace">
            {t >= 1000 ? `${Math.round(t / 1000)}k` : Math.round(t)}
          </text>
        </g>
      ))}
      <path d={area} fill="url(#gAnd)" />
      <path d={linea("contratto")} fill="none" stroke="#D8455F" strokeWidth="2" strokeLinejoin="round" strokeLinecap="round" />
      <path d={linea("spesa")} fill="none" stroke="#8A8A96" strokeWidth="1.5" strokeDasharray="3 3" strokeLinejoin="round" />
      {righe.map((r, i) => (
        <circle key={r.mm} cx={px(i)} cy={py(r.contratto)} r="2.6" fill="#D8455F" />
      ))}
      {righe.map((r, i) => (
        (i === 0 || i === righe.length - 1 || righe.length <= 8) && (
          <text key={`t${r.mm}`} x={px(i)} y={H - 8} textAnchor={i === 0 ? "start" : i === righe.length - 1 ? "end" : "middle"}
            fontSize="10" fill="#6C6C79" fontFamily="JetBrains Mono, ui-monospace, monospace">{meseIt(r.mm).slice(0, 3)}</text>
        )
      ))}
    </svg>
  );
}

/* Distribuzione: un anello, non una torta. Le fette sottili restano leggibili. */
function Anello({ voci, totale, etichetta = "Totale" }) {
  const colori = ["#8E2039", "#D8455F", "#8A8A96", "#4C4C58", "#2E2E38"];
  const R = 54, r = 38, C = 70;
  let acc = 0;
  const archi = voci.map((v, i) => {
    const q = totale ? v.v / totale : 0;
    const a0 = acc * 2 * Math.PI - Math.PI / 2;
    acc += q;
    const a1 = acc * 2 * Math.PI - Math.PI / 2;
    const grande = q > 0.5 ? 1 : 0;
    const p = (ra, a) => [C + ra * Math.cos(a), C + ra * Math.sin(a)];
    const [x0, y0] = p(R, a0), [x1, y1] = p(R, a1), [x2, y2] = p(r, a1), [x3, y3] = p(r, a0);
    return { d: `M${x0} ${y0} A${R} ${R} 0 ${grande} 1 ${x1} ${y1} L${x2} ${y2} A${r} ${r} 0 ${grande} 0 ${x3} ${y3} Z`,
      c: colori[i % colori.length], q, nome: v.k };
  });
  return (
    <div className="riga" style={{ gap: 20, padding: "6px 16px 16px", flexWrap: "wrap" }}>
      <svg viewBox="0 0 140 140" width="140" height="140" role="img" aria-label="Distribuzione della spesa">
        {archi.map((a, i) => <path key={i} d={a.d} fill={a.c} />)}
        <text x="70" y="66" textAnchor="middle" fontSize="15" fill="var(--ink)" fontFamily="JetBrains Mono, ui-monospace, monospace">{eur0(totale)}</text>
        <text x="70" y="82" textAnchor="middle" fontSize="9" fill="#6C6C79" fontFamily="JetBrains Mono, ui-monospace, monospace" letterSpacing="1">{etichetta.toUpperCase()}</text>
      </svg>
      <div style={{ display: "grid", gap: 7, flex: 1, minWidth: 170 }}>
        {archi.map((a, i) => (
          <div className="riga" key={i} style={{ gap: 9, fontSize: 12.5 }}>
            <span style={{ width: 8, height: 8, borderRadius: 3, background: a.c, flex: "none" }} />
            <span>{a.nome}</span>
            <span className="sp" style={{ marginLeft: "auto" }} />
            <span className="mono sub">{pctv(a.q * 100, 0)}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function Conferma({ c, onChiudi }) {
  const [valore, setValore] = useState(c.campo ? c.campo.valore || "" : "");
  return (
    <div className="scrim" onMouseDown={(e) => e.target === e.currentTarget && onChiudi()}>
      <div className="modal" role="dialog" aria-modal="true" aria-label={c.titolo} style={{ maxWidth: 520 }}>
        <header><h2>{c.titolo}</h2></header>
        <div style={{ padding: "16px 18px", whiteSpace: "pre-wrap", fontSize: 13.5, lineHeight: 1.55 }}>
          {c.testo}
          {c.campo && (
            <div className="f" style={{ marginTop: 14 }}>
              <label htmlFor="cf-campo">{c.campo.label}</label>
              <input id="cf-campo" autoFocus value={valore} onChange={(e) => setValore(e.target.value)}
                onKeyDown={(e) => { if (e.key === "Enter" && valore.trim()) { try { c.azione(valore.trim()); } catch (err) { console.error(err); } onChiudi(); } }} />
            </div>
          )}
        </div>
        <footer>
          <span className="sp" />
          <button className="btn ghost" onClick={onChiudi}>Annulla</button>
          {c.azione2 && (
            <button className="btn ghost" onClick={() => { try { c.azione2(); } catch (err) { console.error(err); } onChiudi(); }}>{c.etichetta2}</button>
          )}
          <button className={c.pericolo ? "btn danger" : "btn"}
            disabled={c.campo && !valore.trim()}
            onClick={() => {
              /* Se l'azione fallisce il riquadro si chiude lo stesso: restare bloccati
                 dentro una finestra senza spiegazione è peggio dell'errore. */
              try { c.azione(c.campo ? valore.trim() : true); }
              catch (e) { console.error("Azione non riuscita:", e); }
              onChiudi();
            }}>
            {c.etichetta || "Conferma"}
          </button>
        </footer>
      </div>
    </div>
  );
}

function Editor({ titolo, sotto, fields, draft, setDraft, onSave, onClose, onDelete, liste = {} }) {
  const set = (k, v) => setDraft({ ...draft, [k]: v });
  return (
    <div className="scrim" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" role="dialog" aria-modal="true" aria-label={titolo}>
        <header><h2>{titolo}</h2>{sotto && <p>{sotto}</p>}</header>
        <div className="fields">
          {fields.map((f, i) =>
            f.grp ? (
              <div className="grp" key={"g" + i}>{f.grp}{f.hint && <span>{f.hint}</span>}</div>
            ) : (
              <div className={`f${f.wide ? " wide" : f.w2 ? " w2" : ""}`} key={f.k}>
                <label htmlFor={"f-" + f.k}>{f.label}</label>

                {f.type === "select" ? (
                  <select id={"f-" + f.k} value={draft[f.k] ?? ""} onChange={(e) => set(f.k, e.target.value)}>
                    {(f.opts || []).map((o) => <option key={o.v} value={o.v}>{o.l}</option>)}
                  </select>

                ) : f.type === "textarea" ? (
                  <textarea id={"f-" + f.k} rows={3} value={draft[f.k] ?? ""} onChange={(e) => set(f.k, e.target.value)} />

                ) : f.type === "num" ? (
                  /* Campo calcolato altrove: mostro il valore vero invece di accettare
                     una modifica che poi verrebbe ignorata in silenzio. */
                  f.bloccato !== undefined ? (
                    <input id={"f-" + f.k} className="num" type="text" value={f.bloccato} readOnly aria-readonly="true"
                      tabIndex={-1} style={{ opacity: 0.68, cursor: "not-allowed" }} />
                  ) : (
                    <input id={"f-" + f.k} className="num" type="text" inputMode="decimal" autoComplete="off"
                      value={draft[f.k] ?? ""} placeholder={f.ph || "0"}
                      onChange={(e) => set(f.k, e.target.value.replace(/[^\d.,-]/g, ""))} />
                  )

                ) : f.type === "logo" ? (
                  <div className="logoedit">
                    <Logo src={draft[f.k]} nome={draft.nome} size={48} />
                    <div className="cmd">
                      <input id={"f-" + f.k} type="file" accept="image/png,image/jpeg,image/webp"
                        onChange={(e) => { const fl = e.target.files && e.target.files[0]; if (fl) leggiLogo(fl, (u) => set(f.k, u)); }} />
                      {draft[f.k] && <button className="btn ghost mini" style={{ alignSelf: "flex-start" }} onClick={() => set(f.k, "")}>Rimuovi logo</button>}
                    </div>
                  </div>

                ) : f.type === "time" ? (
                  <input id={"f-" + f.k} type="time" value={draft[f.k] ?? ""} onChange={(e) => set(f.k, e.target.value)} />

                ) : f.type === "colore" ? (
                  <div className="riga" style={{ gap: 8, flexWrap: "wrap" }}>
                    {["#D8455F", "#C99A3A", "#3FA45C", "#4C8DF6", "#8B5CF6", "#E07A3F", "#2BA3A3", "#8A8A96"].map((c) => (
                      <button key={c} aria-label={`Colore ${c}`} onClick={() => set(f.k, c)}
                        style={{ width: 26, height: 26, borderRadius: 8, background: c, flex: "none",
                          boxShadow: draft[f.k] === c ? "0 0 0 2px #fff" : "inset 0 0 0 1px rgba(255,255,255,.2)" }} />
                    ))}
                    <input value={draft[f.k] ?? ""} onChange={(e) => set(f.k, e.target.value)}
                      style={{ width: 110, fontFamily: "var(--mono)", fontSize: 12.5 }} aria-label="Codice colore" />
                  </div>

                ) : f.type === "campagne" ? (
                  <div className="ck">
                    {(() => {
                      const tutte = (liste.campagne || []).filter((k) => !draft.clienteId || k.clienteId === draft.clienteId);
                      const scelte = draft[f.k] || [];
                      const cambia = (id, on) => set(f.k, on ? [...scelte, id] : scelte.filter((x) => x !== id));
                      if (!tutte.length) return <span className="hint">Nessuna campagna per questo cliente. Puoi collegarle più avanti dalla scheda del funnel.</span>;
                      const attive = tutte.filter((k) => ["attiva", "test"].includes(k.stato));
                      return (
                        <>
                          <div className="sceltatesta">
                            <span className="sub">{scelte.length} di {tutte.length} selezionate</span>
                            <span className="sp" />
                            <button className="btn ghost mini" onClick={() => set(f.k, tutte.map((k) => k.id))}>Tutte</button>
                            {attive.length > 0 && attive.length < tutte.length && (
                              <button className="btn ghost mini" onClick={() => set(f.k, attive.map((k) => k.id))}>Solo attive</button>
                            )}
                            <button className="btn ghost mini" disabled={!scelte.length} onClick={() => set(f.k, [])}>Nessuna</button>
                          </div>
                          <div className="sceltarigaintest">
                            <span />
                            <span>Campagna</span>
                            <span>Stato</span>
                            <span>Funnel attuale</span>
                          </div>
                          <div className="sceltalista">
                            {tutte.map((k) => (
                              <label key={k.id} className="sceltariga" data-on={scelte.includes(k.id) ? "1" : "0"}>
                                <input type="checkbox" checked={scelte.includes(k.id)} onChange={(e) => cambia(k.id, e.target.checked)} />
                                <span className="nm">{k.nome}</span>
                                <span className="st">{k.stato}</span>
                                <span className="alt">{k.altro ? `già su ${k.altro}` : ""}</span>
                              </label>
                            ))}
                          </div>
                        </>
                      );
                    })()}
                  </div>

                ) : f.type === "budget" ? (
                  <div className="ck">
                    {(() => {
                      const lista = draft[f.k] || [];
                      const ordinati = [...lista].sort((a, b) => (a.data < b.data ? -1 : 1));
                      const agg = (idx, patch) => set(f.k, lista.map((x, j) => (j === idx ? { ...x, ...patch } : x)));
                      const ultimo = ordinati.slice(-1)[0];
                      return (
                        <>
                          {ordinati.map((b2, i) => {
                            const idx = lista.indexOf(b2);
                            const prec = i > 0 ? parseNumIntl(ordinati[i - 1].importo) : null;
                            const ora = parseNumIntl(b2.importo);
                            return (
                              <div className="pr-riga" key={b2.id} style={{ gridTemplateColumns: "1fr 1fr 1.4fr auto" }}>
                                <input type="date" value={b2.data || ""} onChange={(e) => agg(idx, { data: e.target.value })} aria-label="Dal giorno" />
                                <input className="num" type="text" inputMode="decimal" value={b2.importo ?? ""} placeholder="€/giorno"
                                  onChange={(e) => agg(idx, { importo: e.target.value.replace(/[^\d.,-]/g, "") })} aria-label="Importo" />
                                <input type="text" value={b2.nota || ""} placeholder={prec !== null && ora !== prec ? (ora > prec ? "perché hai alzato" : "perché hai ridotto") : "nota"}
                                  onChange={(e) => agg(idx, { nota: e.target.value })} aria-label="Nota" />
                                <button aria-label="Togli variazione" onClick={() => set(f.k, lista.filter((_, j) => j !== idx))}>×</button>
                              </div>
                            );
                          })}
                          <button className="btn ghost mini" style={{ alignSelf: "flex-start" }}
                            onClick={() => set(f.k, [...lista, { id: uid(), data: oggi(), importo: ultimo ? ultimo.importo : (draft.budgetGiorno || ""), nota: "" }])}>
                            + Nuovo budget giornaliero
                          </button>
                          <span className="hint">
                            Ogni riga vale dal giorno indicato fino alla successiva. L'erogazione attesa tiene conto di tutti i cambi,
                            quindi il confronto con lo speso resta corretto anche se il budget cambia più volte.
                          </span>
                        </>
                      );
                    })()}
                  </div>

                ) : f.type === "fasi" ? (
                  <div className="ck">
                    {(() => {
                      const lista = draft[f.k] || [];
                      const ordina = (l) => [...l].sort((a, b) => (a.inizio < b.inizio ? -1 : 1));
                      const agg = (idx, patch) => set(f.k, lista.map((x, j) => (j === idx ? { ...x, ...patch } : x)));
                      const nSett = (pz) => lista.filter((x) => x.tipo === "attiva" && x.inizio <= pz.inizio).length;
                      const ultima = ordina(lista).slice(-1)[0];
                      const prossimoDal = ultima ? sommaGiorni(ultima.fine, 1) : (draft.inizio ? sommaGiorni(draft.inizio, TEST_2) : oggi());
                      return (
                        <>
                          {ordina(lista).map((pz) => {
                            const idx = lista.indexOf(pz);
                            return (
                              <div className="pr-riga" key={pz.id} style={{ gridTemplateColumns: "1.1fr 1fr 1fr auto" }}>
                                <span className="sub" style={{ fontWeight: 600 }}>
                                  {pz.tipo === "attiva" ? `Attiva · sett. ${nSett(pz)}` : et(TIPI_PERIODO, pz.tipo)}
                                </span>
                                <input type="date" value={pz.inizio || ""} onChange={(e) => agg(idx, { inizio: e.target.value })} aria-label="Dal" />
                                <input type="date" value={pz.fine || ""} onChange={(e) => agg(idx, { fine: e.target.value })} aria-label="Al" />
                                <button aria-label="Togli fase" onClick={() => set(f.k, lista.filter((_, j) => j !== idx))}>×</button>
                              </div>
                            );
                          })}
                          <div className="riga" style={{ gap: 8, flexWrap: "wrap" }}>
                            {!lista.some((x) => x.tipo === "test") && draft.inizio && (
                              <button className="btn ghost mini" onClick={() => set(f.k, [...lista,
                                { id: uid(), tipo: "test", inizio: draft.inizio, fine: sommaGiorni(draft.inizio, TEST_1 - 1) }])}>
                                + Fase Test ({TEST_1} giorni)
                              </button>
                            )}
                            {!lista.some((x) => x.tipo === "matura") && draft.inizio && (
                              <button className="btn ghost mini" onClick={() => set(f.k, [...lista,
                                { id: uid(), tipo: "matura", inizio: sommaGiorni(draft.inizio, TEST_1), fine: sommaGiorni(draft.inizio, TEST_2 - 1) }])}>
                                + Fase Matura ({TEST_2 - TEST_1} giorni)
                              </button>
                            )}
                            <button className="btn ghost mini" onClick={() => set(f.k, [...lista,
                              { id: uid(), tipo: "attiva", inizio: prossimoDal, fine: sommaGiorni(prossimoDal, FINESTRA - 1) }])}>
                              + Settimana successiva
                            </button>
                          </div>
                          <span className="hint">
                            Le date le tengo io in fila: ogni settimana parte dal giorno dopo la fine della precedente. I numeri di ciascuna finestra si inseriscono dal pulsante Fasi nella tabella.
                          </span>
                        </>
                      );
                    })()}
                  </div>

                ) : f.type === "flag" ? (
                  <label className="ckr" style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}>
                    <input id={"f-" + f.k} type="checkbox" checked={draft[f.k] === true}
                      onChange={(e) => set(f.k, e.target.checked)} />
                    <span>{f.hint || "Sì"}</span>
                  </label>
                ) : f.type === "check" ? (
                  <div className="ck">
                    {(draft[f.k] || []).map((r, idx) => (
                      <div className="ckr" key={r.id}>
                        <input type="checkbox" checked={!!r.fatto} aria-label="Passaggio completato"
                          onChange={(e) => set(f.k, draft[f.k].map((x, j) => (j === idx ? { ...x, fatto: e.target.checked } : x)))} />
                        <input type="text" value={r.t} placeholder="Passaggio"
                          onChange={(e) => set(f.k, draft[f.k].map((x, j) => (j === idx ? { ...x, t: e.target.value } : x)))} />
                        <button aria-label="Togli passaggio" onClick={() => set(f.k, draft[f.k].filter((_, j) => j !== idx))}>×</button>
                      </div>
                    ))}
                    <button className="btn ghost mini" style={{ alignSelf: "flex-start" }}
                      onClick={() => set(f.k, [...(draft[f.k] || []), { id: uid(), t: "", fatto: false }])}>+ Passaggio</button>
                  </div>

                ) : f.type === "passi" ? (
                  <div className="ck">
                    {(draft[f.k] || []).map((r, idx) => {
                      const agg = (patch) => set(f.k, draft[f.k].map((x, j) => (j === idx ? { ...x, ...patch } : x)));
                      return (
                        <div className="pr-riga" key={r.id}>
                          <input type="text" value={r.nome} placeholder="Nome del passo" onChange={(e) => agg({ nome: e.target.value })} />
                          <select value={r.tipo} onChange={(e) => agg({ tipo: e.target.value })} aria-label="Tipo di passo">
                            {TIPI_PASSO.map((t) => <option key={t.v} value={t.v}>{t.l}</option>)}
                          </select>
                          <input className="num" type="text" inputMode="decimal" value={r.visite ?? ""} placeholder="visite"
                            onChange={(e) => agg({ visite: e.target.value.replace(/[^\d.,]/g, "") })} />
                          <input className="num" type="text" inputMode="decimal" value={r.conversioni ?? ""} placeholder="passano"
                            onChange={(e) => agg({ conversioni: e.target.value.replace(/[^\d.,]/g, "") })} />
                          <button aria-label="Togli passo" onClick={() => set(f.k, draft[f.k].filter((_, j) => j !== idx))}>×</button>
                        </div>
                      );
                    })}
                    <button className="btn ghost mini" style={{ alignSelf: "flex-start" }}
                      onClick={() => set(f.k, [...(draft[f.k] || []), { id: uid(), nome: "", tipo: "pagina", visite: "", conversioni: "" }])}>+ Passo</button>
                  </div>

                ) : (
                  <>
                    <input id={"f-" + f.k} type={f.type || "text"} value={draft[f.k] ?? ""}
                      list={f.lista ? `dl-${f.lista}` : undefined}
                      onChange={(e) => set(f.k, e.target.value)} />
                    {f.lista && (liste[f.lista] || []).length > 0 && (
                      <datalist id={`dl-${f.lista}`}>
                        {liste[f.lista].map((x) => <option key={x} value={x} />)}
                      </datalist>
                    )}
                  </>
                )}

                {f.hint && f.type !== "logo" && <span className="hint">{f.hint}</span>}
              </div>
            )
          )}
        </div>
        <footer>
          {onDelete && <button className="btn danger" onClick={onDelete}>Elimina</button>}
          <span className="sp" />
          <button className="btn ghost" onClick={onClose}>Annulla</button>
          <button className="btn" onClick={onSave}>Salva</button>
        </footer>
      </div>
    </div>
  );
}

const NUMERICI = {
  agenzia: ["costoOrario"], aree: [], clienti: ["fee", "budgetMese", "margineLordo", "resi", "aov"], contatti: [], trattative: ["valore"],
  campagne: ["budgetGiorno", "budget", "spesa", "impression", "click", "clickLink", "lead", "conversioni", "ricavi", "valoreContratto", "incassato", "frequenza", "roasMin", "cpaMax", "cplMax", "ctrMin", "freqMax", "cpmMax"],
  task: ["stima"],
  vendite: [...CAMPI_VENDITA, "slot", "oreContatto", "convPiattaforma"],
  chiusure: ["valoreContratto", "incassato"],
  automazioni: ["valore", "giorno", "taskGiorni"],
  creativi: ["spesa", "impression", "click", "clickLink", "lead", "conversioni", "ricavi", "valoreContratto", "incassato", "frequenza"],
  periodi: ["spesa", "impression", "click", "clickLink", "lead", "conversioni", "ricavi", "valoreContratto", "incassato", "frequenza"],
  annotazioni: [],
  esperimenti: [],
  funnel: ["spesa", "valoreMedio"],
  file: [],
  messaggio: [],
};

/* ================================================================== */
/* App                                                                 */
/* ================================================================== */

/* onEsci e utente arrivano dalla schermata di accesso (App.jsx); senza onEsci il tasto non compare */
/* meta: le funzioni per Meta Ads (src/lib/metaApi.js); senza, la sezione Meta spiega che serve la versione online */
/* Catalogo di esempio dell'Academy: si vede finché non ci sono i corsi veri (data.academy.corsi).
   sezione = dove il CRM suggerisce la lezione (advertising, vendite, funnel, winner). */
const LOGO_ACADEMY = "/academy-logo.webp";
/* Macro sezioni dell'Accademia, nell'ordine della piattaforma */
const ACADEMY_MACRO = [
  { v: "corsi", l: "Corsi", d: "I percorsi completi, dalle basi allo scaling." },
  { v: "formazione", l: "Formazione Specifica", d: "Approfondimenti verticali su un argomento alla volta." },
  { v: "guide", l: "Guide Strategiche", d: "Metodi e checklist da applicare subito ai clienti." },
  { v: "live", l: "Eventi Live", d: "Dirette, masterclass e sessioni di domande e risposte." },
];
/* Catalogo finché la piattaforma non è collegata (data.academy.corsi vuoto): i 5 corsi della
   sezione Corsi hanno titolo e copertina veri; descrizioni, lezioni e le altre sezioni sono esempi.
   sezione = dove il CRM suggerisce il corso o la lezione (advertising, winner, funnel, vendite). */
const ACADEMY_ESEMPIO = [
  { id: "c-fb-principianti", macro: "corsi", titolo: "Facebook Ads per Principianti", copertina: "/academy/facebook-ads-principianti.webp", colore: "#3B5A9A", livello: "Base", sezione: "advertising",
    descrizione: "Le basi per partire con le campagne su Facebook e Instagram: struttura, pubblici, budget e lettura dei risultati.", lezioni: [] },
  { id: "c-bootcamp", macro: "corsi", titolo: "Facebook Ads Bootcamp", copertina: "/academy/facebook-ads-bootcamp.webp", colore: "#8E2039", livello: "Intermedio", sezione: "advertising",
    descrizione: "Il percorso intensivo per impostare, testare e ottimizzare campagne che portano risultati.", lezioni: [] },
  { id: "c-scalare", macro: "corsi", titolo: "Come Convertire e Scalare il Traffico a Pagamento", copertina: "/academy/convertire-scalare-traffico.webp", colore: "#3A3A3A", livello: "Avanzato", sezione: "winner",
    descrizione: "Trasformare il traffico in clienti e aumentare la spesa mantenendo sotto controllo il costo per acquisizione.", lezioni: [] },
  { id: "c-landing", macro: "corsi", titolo: "Come Creare ed Ottimizzare Landing Page e Web Page ad Alta Conversione", copertina: "/academy/landing-page-alta-conversione.webp", colore: "#1B4A5A", livello: "Intermedio", sezione: "funnel",
    descrizione: "Costruire pagine che convertono il traffico delle campagne in contatti e vendite.", lezioni: [] },
  { id: "c-ios14", macro: "corsi", titolo: "iOS 14 Senza Segreti", copertina: "/academy/ios14-senza-segreti.webp", colore: "#2A2A2A", livello: "Intermedio", sezione: "advertising",
    descrizione: "Cosa è cambiato con iOS 14 nel tracciamento e come impostare eventi e misurazione oggi.", lezioni: [] },
  { id: "es-pixel", macro: "formazione", esempio: true, titolo: "Pixel e tracciamento avanzato", colore: "#5C1B2C", livello: "Avanzato", gratuito: true, sezione: "advertising",
    descrizione: "Esempio di formazione specifica: eventi, API di conversione e verifica dei dati.",
    lezioni: [{ id: "es-pixel-1", titolo: "Eventi standard e personalizzati", durata: "12 min" }, { id: "es-pixel-2", titolo: "API di conversione", durata: "15 min" }] },
  { id: "es-copy", macro: "formazione", esempio: true, titolo: "Copy per le inserzioni", colore: "#B4566B", livello: "Intermedio", prezzo: "97 €",
    descrizione: "Esempio di formazione specifica: hook, angoli e testi che fermano lo scroll.",
    lezioni: [{ id: "es-copy-1", titolo: "Anatomia di un hook", durata: "9 min", gratuita: true, sezione: "winner" }, { id: "es-copy-2", titolo: "Dieci angoli da testare", durata: "14 min" }] },
  { id: "es-lancio", macro: "guide", esempio: true, titolo: "Checklist di lancio campagna", colore: "#3E1720", livello: "Base", gratuito: true,
    descrizione: "Esempio di guida strategica: i controlli da fare prima di accendere una campagna.",
    lezioni: [{ id: "es-lancio-1", titolo: "Prima del lancio", durata: "7 min" }, { id: "es-lancio-2", titolo: "Le prime 72 ore", durata: "8 min", sezione: "advertising" }] },
  { id: "es-bench", macro: "guide", esempio: true, titolo: "Benchmark per settore", colore: "#6E1E31", livello: "Intermedio", prezzo: "47 €",
    descrizione: "Esempio di guida strategica: costi per lead, appuntamento e cliente di riferimento per settore.",
    lezioni: [{ id: "es-bench-1", titolo: "Come leggere i benchmark", durata: "10 min", gratuita: true, sezione: "winner" }] },
  { id: "es-live-analisi", macro: "live", esempio: true, titolo: "Live mensile: analisi degli account", colore: "#8E2039", livello: "Tutti",
    descrizione: "Esempio di evento live: analisi in diretta degli account degli iscritti.",
    lezioni: [{ id: "es-live-1", titolo: "Registrazione di ottobre", durata: "62 min" }] },
  { id: "es-live-qa", macro: "live", esempio: true, titolo: "Domande e risposte con il team", colore: "#4A1B26", livello: "Tutti", gratuito: true,
    descrizione: "Esempio di evento live: sessione aperta di domande e risposte.",
    lezioni: [{ id: "es-live-qa-1", titolo: "Q&A di settembre", durata: "48 min" }] },
];

export default function CRM({ onEsci, utente = "", meta = null, telegram = null, integrazioni = null, analisi = null, spazio = null } = {}) {
  const [data, setData] = useState(null);
  /* Aspetto: scuro (predefinito) o chiaro, scelto in Impostazioni e ricordato da questo browser */
  const [tema, setTema] = useState(() => { try { return localStorage.getItem("crm-tema") === "chiaro" ? "chiaro" : "scuro"; } catch { return "scuro"; } });
  const cambiaTema = (t) => { setTema(t); try { localStorage.setItem("crm-tema", t); } catch { /* browser senza archivio: vale per questa sessione */ } };
  const [cliAtt, setCliAtt] = useState("");
  const [vista, setVista] = useState("cruscotto");
  const [q, setQ] = useState("");
  const [ed, setEd] = useState(null);
  const [salvato, setSalvato] = useState(false);
  const [vistaTask, setVistaTask] = useState("bacheca");
  const [filtri, setFiltri] = useState({ area: "", prio: "", cli: "", fatte: false });
  const [compatto, setCompatto] = useState(() => typeof window !== "undefined" && window.innerWidth < 900);
  const [graf, setGraf] = useState({ ambito: "campagne", gruppo: "campagna", metrica: "spesa", tipo: "barre" });
  const [fAdv, setFAdv] = useState({ cli: "", piattaforma: "", obiettivo: "", stato: "", funnel: "", linea: "", faseCorrente: "", segnale: "" });
  const [piano, setPiano] = useState({ obiettivo: "", cpl: "" });
  const [funnelAperto, setFunnelAperto] = useState("");
  const [daRipristinare, setDaRipristinare] = useState(null);
  const [imp, setImp] = useState({ tipo: "campagne", clienteId: "", profilo: "", url: "", testo: "", nomeFile: "", intestazioni: [], righe: [], mappa: {}, esito: null, fogli: [], foglio: "" });
  const [capitolo, setCapitolo] = useState("inizio");
  const [chat, setChat] = useState({ messaggi: [], testo: "", attesa: false, errore: "", ambito: "cliente" });
  const [fFile, setFFile] = useState({ cat: "", tipo: "" });
  const [fileAperto, setFileAperto] = useState("");
  const [conferma, setConferma] = useState(null);
  const [consolidato, setConsolidato] = useState(true);
  const [mostraArch, setMostraArch] = useState(false);
  const [ordAdv, setOrdAdv] = useState("recenti");
  const [richiesta, setRichiesta] = useState(null);

  /* Traduzione: chiave mancante o lingua incompleta ricadono sull'italiano,
     così una traduzione parziale non lascia buchi al posto delle parole. */
  /* I dati arrivano dopo il primo disegno: prima di allora la lingua è l'italiano */
  const lingua = (data && data.agenzia && data.agenzia.lingua) || "it";
  impostaLingua(lingua);
  /* La valuta è quella del cliente aperto (o della madre), o quella dell'agenzia se non ne ha una propria */
  impostaValuta(cliAtt ? valutaCliente(data, cliAtt) : ((data && data.agenzia && data.agenzia.valuta) || "EUR"));
  const t = (chiave, prede) => (DIZ[base(lingua)] && DIZ[base(lingua)][chiave]) || DIZ.it[chiave] || prede || chiave;
  const [cruOrd, setCruOrd] = useState("recenti");
  const [cruStato, setCruStato] = useState("");
  const [cruFase, setCruFase] = useState("");
  const [cruSegnale, setCruSegnale] = useState("");
  const [pxEvento, setPxEvento] = useState("Lead");
  const [pxVal, setPxVal] = useState({});
  const [pxCopiato, setPxCopiato] = useState(false);
  const [adaCliente, setAdaCliente] = useState("");
  const [semCliente, setSemCliente] = useState("");
  const [semCamp, setSemCamp] = useState([]);
  const [calVista, setCalVista] = useState("mese");
  const [calData, setCalData] = useState("");
  const [calGiorno, setCalGiorno] = useState("");
  const [calFonti, setCalFonti] = useState({ eventi: true, task: true, campagne: true, finestre: false, chiusure: true });
  const [filtriAperti, setFiltriAperti] = useState(false);
  const [mostraArchCli, setMostraArchCli] = useState(false);
  const [trascina, setTrascina] = useState(false);
  const [apertiCli, setApertiCli] = useState({});
  /* Advertising: sottoclienti aperti sotto il loro cliente principale */
  const [apertiAdv, setApertiAdv] = useState({});
  const [campFasi, setCampFasi] = useState("");
  const [fase, setFase] = useState("totale");
  const [nuovoPadre, setNuovoPadre] = useState("");
  const [allegato, setAllegato] = useState(null);
  const [estrazione, setEstrazione] = useState(null);
  const chiedi = (o) => setConferma(o);
  const [erroreArchivio, setErroreArchivio] = useState("");
  const [erroreSalvataggio, setErroreSalvataggio] = useState("");
  const bloccaSalvataggio = useRef(false);
  /* Spazio di lavoro appena creato (nessun archivio salvato): serve a proporre il fuso orario */
  const primoAvvio = useRef(false);
  /* Dopo un ripristino o un reset voluto si torna a salvare normalmente */
  const sbloccaArchivio = () => { bloccaSalvataggio.current = false; setErroreArchivio(""); };

  useEffect(() => {
    (async () => {
      let d = null, guasto = "";
      /* Distinguo "non c'è niente" da "non riesco a leggere": nel secondo caso NON carico
         i dati di esempio e NON salvo, altrimenti sovrascriverei l'archivio vero. */
      const leggi = async (chiave) => {
        let r;
        try { r = await window.storage.get(chiave); } catch (e) {
          /* Alcuni ambienti segnalano con un errore anche la chiave che non esiste:
             controllo l'elenco per capire se l'archivio c'è davvero ma non si legge. */
          let esiste = false;
          try {
            const l = await window.storage.list("crm:");
            const chiavi = ((l && l.keys) || []).map((x) => (typeof x === "string" ? x : x && x.key));
            esiste = chiavi.includes(chiave);
          } catch { esiste = false; }
          if (!esiste) return null;
          throw new Error("lettura");
        }
        if (!r || !r.value) return null;
        try { return JSON.parse(r.value); } catch { throw new Error("formato"); }
      };
      try { d = await leggi(KEY); } catch (e) { guasto = e.message; }
      if (!d && !guasto) { try { const v = await leggi(KEY_V5); if (v) d = migra(v); } catch (e) { guasto = e.message; } }
      if (guasto) {
        bloccaSalvataggio.current = true;
        setErroreArchivio(guasto === "formato"
          ? "L'archivio salvato risulta danneggiato. Non ho sovrascritto nulla: ripristina un backup dalla sezione Backup."
          : "Non riesco a leggere l'archivio salvato. Non ho sovrascritto nulla: ricarica la pagina; se persiste, ripristina un backup.");
      }
      /* Anche i dati di esempio passano da normalizza: le raccolte aggiunte dopo
         (giornate, eventi, categorie, modelli, periodi…) non sono nel SEED, e
         senza questo passaggio il primo avvio trova collezioni inesistenti. */
      primoAvvio.current = !d && !guasto;
      setData(normalizza(d || SEED()));
    })();
  }, []);

  useEffect(() => {
    if (!data || bloccaSalvataggio.current) return;
    let vivo = true;
    const t = setTimeout(async () => {
      try {
        await window.storage.set(KEY, JSON.stringify(data));
        if (!vivo) return;
        setErroreSalvataggio("");
        setSalvato(true);
        setTimeout(() => vivo && setSalvato(false), 1400);
      } catch (e) {
        console.error("Salvataggio non riuscito", e);
        if (vivo) setErroreSalvataggio("Le ultime modifiche NON sono state salvate (spazio pieno o archivio non disponibile). Scarica subito un backup dalla sezione Backup.");
      }
    }, 400);
    return () => { vivo = false; clearTimeout(t); };
  }, [data]);

  /* Copia automatica giornaliera nell'archivio interno del browser.
     Non sostituisce il file scaricato: protegge dai dati corrotti, non dal browser perso. */
  useEffect(() => {
    if (!data || !data.backup || data.backup.auto !== "si" || bloccaSalvataggio.current) return;
    if (data.backup.ultimoAuto === oggi()) return;
    let vivo = true;
    (async () => {
      const chiave = `crm:snap:${oggi()}`;
      const testo = JSON.stringify({ ...data, aiKey: "" });
      try {
        await window.storage.set(chiave, testo);
        if (!vivo) return;
        setData((d) => {
          const nuove = [{ chiave, quando: new Date().toISOString(), peso: testo.length }, ...(d.snapshots || []).filter((s) => s.chiave !== chiave)];
          nuove.slice(6).forEach((s) => { try { window.storage.delete(s.chiave); } catch (e) { /* già rimossa */ } });
          return { ...d, snapshots: nuove.slice(0, 6), backup: { ...d.backup, ultimoAuto: oggi(), errore: "" } };
        });
      } catch (e) {
        console.error("Copia automatica non riuscita", e);
        if (!vivo) return;
        setData((d) => ({ ...d, backup: { ...d.backup, ultimoAuto: oggi(), errore: "Spazio insufficiente: libera copie vecchie o riduci i loghi caricati." } }));
      }
    })();
    return () => { vivo = false; };
  }, [data]);

  const upd = useCallback((coll, fn) => setData((d) => ({ ...d, [coll]: fn(d[coll]) })), []);

  /* Motore delle automazioni: gira a ogni cambiamento, mentre l'app è aperta */
  useEffect(() => {
    if (!data || !data.automazioni || !(data.automazioni || []).some((a) => a.attiva === "si")) return;
    /* Stessi numeri delle schermate: totali dalle fasi, partenza reale, niente clienti archiviati */
    const campagne = campagneCalcolate(data).filter((k) => k.archiviata !== "si");
    const eventi = valutaAutomazioni(data, campagne);
    if (!eventi.length) return;

    setData((d) => {
      let task = [...d.task];
      let registro = [...(d.registro || [])];
      const eseguiti = [...(d.eseguiti || [])];
      eventi.forEach((e) => {
        eseguiti.push(e.chiave);
        const a = e.aut;
        /* Un task identico ancora aperto basta: non se ne crea uno nuovo ogni giorno */
        const giaAperto = task.some((t) => t.autoRif === e.rif && t.stato !== "fatto");
        if (a.azTask === "si" && !giaAperto) {
          task = [{
            id: uid(), autoRif: e.rif, clienteId: e.clienteId || a.clienteId || "",
            areaId: a.taskArea || (d.aree[0] || {}).id || "", titolo: a.taskTitolo || e.testo,
            note: `${e.testo}\n${e.dettaglio}\n\n(creato dall'automazione "${a.nome}")`,
            stato: "dafare", priorita: a.taskPriorita || "media", inizio: oggi(),
            scadenza: piu(a.taskGiorni || 2), stima: 0, ricorrenza: "no", checklist: [], creato: oggi(), auto: true,
          }, ...task];
        }
        registro = [{
          id: uid(), quando: new Date().toISOString(), automazione: a.nome, testo: e.testo,
          esito: [a.azTask === "si" ? (giaAperto ? "task già aperto" : "task creato") : null, a.azWebhook === "si" ? (a.url ? "webhook inviato" : "webhook senza URL") : null]
            .filter(Boolean).join(" · ") || "solo registrato",
        }, ...registro];
      });
      return { ...d, task, registro: registro.slice(0, 80), eseguiti: eseguiti.slice(-5000) };
    });

    eventi.filter((e) => e.aut.azWebhook === "si" && e.aut.url).forEach((e) => {
      inviaWebhook(e.aut.url, {
        automazione: e.aut.nome, innesco: e.aut.trigger, evento: e.testo,
        dettaglio: e.dettaglio, dati: e.contesto, quando: new Date().toISOString(),
      });
    });
  }, [data]);

  /* ---------- Meta Ads ----------
     I numeri arrivano dalla funzione meta-sync (tabelle meta_*): qui li leggo all'apertura e
     dopo ogni aggiornamento, e li applico a campagne, ad set, annunci e finestre.
     Quale account appartiene a quale cliente lo decidi tu, dalla scheda o dalle Impostazioni. */
  const [metaInfo, setMetaInfo] = useState(null);
  const [metaStato, setMetaStato] = useState({ attesa: "", errore: "", esito: "" });
  const [metaScelta, setMetaScelta] = useState({});
  /* Scelta delle campagne di un account: { account, lista (da Meta), scelte, filtro, errore } */
  const [metaRegola, setMetaRegola] = useState(null);
  /* Import di un file KPI nel Ciclo di vendita: { nome, tabelle, scelte, cliente, anno, errore } */
  const [kpiImport, setKpiImport] = useState(null);
  const metaAvviato = useRef(false);

  /* ---------- Bot Telegram: comandi in coda (task, note, modifiche) applicati dal CRM aperto ---------- */
  const [tgChat, setTgChat] = useState([]);
  const [tgStato, setTgStato] = useState({ attesa: "", errore: "", esito: "", link: "", applicati: 0 });
  const caricaTelegram = useCallback(async () => {
    if (!telegram) return;
    try { setTgChat(await telegram.chat()); } catch { /* prima della migrazione 0006 */ }
  }, [telegram]);
  const datiPronti = !!data;
  useEffect(() => {
    if (!telegram || !datiPronti) return;
    let vivo = true;
    const giro = async () => {
      if (!vivo || bloccaSalvataggio.current) return;
      try {
        const comandi = await telegram.prendiComandi();
        if (!vivo || !comandi.length) return;
        setData((d) => (d ? applicaComandi(d, comandi, { uid, oggi: oggi(), giorniRicorrenza: (v) => (RICORRENZE.find((r) => r.v === v) || {}).g || 0 }).data : d));
        setTgStato((x) => ({ ...x, applicati: x.applicati + comandi.length, ultimo: new Date().toISOString() }));
      } catch { /* riprovo al prossimo giro */ }
    };
    caricaTelegram();
    giro();
    const t = setInterval(giro, 60000);
    return () => { vivo = false; clearInterval(t); };
  }, [telegram, datiPronti, caricaTelegram]);
  const azioneTelegram = async (fn, ok) => {
    setTgStato((x) => ({ ...x, attesa: "1", errore: "", esito: "" }));
    try {
      const r = await fn();
      await caricaTelegram();
      setTgStato((x) => ({ ...x, attesa: "", ...(typeof ok === "function" ? ok(r) : { esito: ok || "" }) }));
    } catch (e) { setTgStato((x) => ({ ...x, attesa: "", errore: e.message || String(e) })); }
  };
  /* ---------- Software esterni (Boschetto): le task arrivano dalla coda dei comandi;
     qui il ritorno, cioè i cambi fatti nel CRM che tornano al software di origine ---------- */
  const [intLista, setIntLista] = useState([]);
  const [intStato, setIntStato] = useState({ attesa: "", errore: "", esito: "", chiave: null, indirizzo: "", nome: "Boschetto", base: "" });
  const intSospese = useRef({});
  const caricaIntegrazioni = useCallback(async () => {
    if (!integrazioni) return;
    try {
      const r = await integrazioni.stato(), lista = r.integrazioni || [];
      setIntLista(lista); setIntStato((x) => ({ ...x, indirizzo: r.indirizzo || "" }));
      /* Ogni integrazione del server ha la sua scheda nei dati (cliente principale e abbinamenti) */
      const task = lista.filter((x) => x.tipo !== "mcp");
      setData((d) => (!d || task.every((x) => (d.integrazioni || []).some((l) => l.id === x.id)) ? d
        : { ...d, integrazioni: [...(d.integrazioni || []), ...task.filter((x) => !(d.integrazioni || []).some((l) => l.id === x.id)).map((x) => ({ id: x.id, nome: x.nome, clienteBase: "", abbinamenti: {}, nomiVisti: {} }))] }));
    }
    catch { /* funzione non ancora pubblicata */ }
  }, [integrazioni]);
  useEffect(() => { if (integrazioni && datiPronti) caricaIntegrazioni(); }, [integrazioni, datiPronti, caricaIntegrazioni]);
  const taskEsterne = data ? data.task.filter((t) => t.esterno && !t.esterno.eliminata) : [];
  const firmaRitorno = taskEsterne.map((t) => t.id + JSON.stringify(statoEsterno(t)) + JSON.stringify(t.esterno.ultimo)).join("|");
  useEffect(() => {
    if (!integrazioni || !firmaRitorno) return;
    const t = setTimeout(async () => {
      const gruppi = new Map();
      taskEsterne.forEach((x) => {
        const ora = statoEsterno(x), prima = x.esterno.ultimo || {};
        const cambiati = Object.keys(ora).filter((k) => ora[k] !== (prima[k] || ""));
        if (!cambiati.length || intSospese.current[x.esterno.integrazioneId]) return;
        const g = gruppi.get(x.esterno.integrazioneId) || [];
        g.push({ task: x.id, ora, m: { esternoId: x.esterno.id, ...ora, cambiati } });
        gruppi.set(x.esterno.integrazioneId, g);
      });
      for (const [id, lista] of gruppi) {
        try {
          await integrazioni.notifica(id, lista.map((x) => x.m));
          const inviati = new Map(lista.map((x) => [x.task, x.ora]));
          setData((d) => (d ? { ...d, task: d.task.map((x) => (inviati.has(x.id) && x.esterno ? { ...x, esterno: { ...x.esterno, ultimo: inviati.get(x.id) } } : x)) } : d));
        } catch (e) {
          /* Non riprovo a ogni modifica: fino al prossimo caricamento della pagina resta in sospeso */
          intSospese.current[id] = true;
          setIntStato((x) => ({ ...x, errore: `Non riesco a mandare i cambi al software esterno: ${e.message || e}` }));
        }
      }
    }, 3000);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [integrazioni, firmaRitorno]);
  const azioneIntegrazione = async (fn, ok) => {
    setIntStato((x) => ({ ...x, attesa: "1", errore: "", esito: "" }));
    try {
      const r = await fn();
      await caricaIntegrazioni();
      setIntStato((x) => ({ ...x, attesa: "", ...(typeof ok === "function" ? ok(r) : { esito: ok || "" }) }));
      return r;
    } catch (e) { setIntStato((x) => ({ ...x, attesa: "", errore: e.message || String(e) })); return null; }
  };
  /* ---------- Winner ads: numeri degli annunci e analisi salvati dal connettore per Claude ---------- */
  const [winDati, setWinDati] = useState({ chiave: "", ads: [], testi: [], carico: false, errore: "" });
  const [winFoto, setWinFoto] = useState("");
  /* Academy: filtro del catalogo, corso aperto, lezione aperta */
  const [acFiltro, setAcFiltro] = useState("tutti");
  const [acCorso, setAcCorso] = useState("");
  const [acLezione, setAcLezione] = useState(null);
  useEffect(() => {
    if (!acCorso) return;
    const esc = (e) => { if (e.key === "Escape") { setAcLezione(null); setAcCorso(""); } };
    window.addEventListener("keydown", esc);
    return () => window.removeEventListener("keydown", esc);
  }, [acCorso]);
  const [winGiro, setWinGiro] = useState(0);
  const winIds = data && cliAtt ? (consolidato
    ? [cliAtt, ...(function sotto(id) { return data.clienti.filter((z) => z.padreId === id && z.archiviato !== "si").flatMap((z) => [z.id, ...sotto(z.id)]); })(cliAtt)]
    : [cliAtt]) : [];
  const winChiave = vista === "winner" ? winIds.join(",") + "#" + winGiro : "";
  useEffect(() => {
    if (!analisi || !winChiave) return;
    let vivo = true;
    const ids = winChiave.split("#")[0].split(",").filter(Boolean);
    setWinDati((x) => ({ ...x, carico: true, errore: "" }));
    Promise.all([analisi.ads(ids), analisi.testi(ids)])
      .then(([ads, testi]) => { if (vivo) setWinDati({ chiave: winChiave, ads, testi, carico: false, errore: "" }); })
      .catch((e) => { if (vivo) setWinDati((x) => ({ ...x, carico: false, errore: e.message || String(e) })); });
    return () => { vivo = false; };
  }, [analisi, winChiave]);
  /* ---------- Fuso orario e ora dell'aggiornamento di Meta (migrazione 0010) ---------- */
  const [orari, setOrari] = useState({ fuso: "Europe/Rome", ora_meta: 9, impostati: false, errore: "", esito: "" });
  useEffect(() => {
    if (!spazio || !datiPronti) return;
    let vivo = true;
    (async () => {
      try {
        const r = await spazio.leggi();
        if (!vivo) return;
        if (r) { setOrari((x) => ({ ...x, fuso: r.fuso, ora_meta: r.ora_meta, impostati: true })); return; }
        // Spazio nuovo: il fuso del computer di chi lo crea; gli spazi esistenti restano su Roma alle 9
        if (primoAvvio.current) {
          const fuso = fusoDelBrowser();
          await spazio.salva({ fuso, ora_meta: 9 });
          if (vivo) setOrari((x) => ({ ...x, fuso, ora_meta: 9, impostati: true }));
        }
      } catch { /* prima della migrazione 0010: valgono Roma e le 9 */ }
    })();
    return () => { vivo = false; };
  }, [spazio, datiPronti]);
  const salvaOrari = async (campi) => {
    const prima = orari;
    setOrari((x) => ({ ...x, ...campi, errore: "", esito: "" }));
    try { await spazio.salva({ fuso: campi.fuso ?? prima.fuso, ora_meta: campi.ora_meta ?? prima.ora_meta }); setOrari((x) => ({ ...x, impostati: true, esito: "Salvato." })); }
    catch (e) { setOrari({ ...prima, errore: `Non riesco a salvare: ${e.message || e}` }); }
  };
  const caricaMeta = useCallback(async (conDati = true, seguita = null) => {
    if (!meta) return null;
    try {
      const info = await meta.stato(seguita);
      setMetaInfo(info);
      if (conDati && info.collegamenti.length && !bloccaSalvataggio.current) {
        const dati = await meta.dati();
        setData((d) => {
          if (!d) return d;
          const r = applicaMeta(d, { ...info, ...dati }, {
            oggi: oggi(), uid, bmkDi: (o) => BMK_DEFAULT[o] || BMK_VUOTO, valutaAgenzia: (d.agenzia && d.agenzia.valuta) || "EUR",
          });
          return r.cambiato ? r.data : d;
        });
      }
      return info;
    } catch (e) {
      setMetaStato((x) => ({ ...x, errore: `Non riesco a leggere i dati di Meta: ${e.message || e}` }));
      return null;
    }
  }, [meta]);
  useEffect(() => {
    if (!meta || !data || metaAvviato.current) return;
    metaAvviato.current = true;
    (async () => {
      const info = await caricaMeta(true);
      if (!info) return;
      /* Il vecchio campo "Account pubblicitario": se l'avevi compilato diventa un collegamento */
      const presi = new Set(info.collegamenti.map((c) => String(c.account_id)));
      const daFare = data.clienti.filter((z) => z.accountAdv && z.archiviato !== "si")
        .map((z) => ({ z, id: String(z.accountAdv).replace(/^act_/i, "").trim() }))
        .filter((x) => /^\d+$/.test(x.id) && !presi.has(x.id));
      if (!daFare.length) return;
      try { for (const x of daFare) await meta.collega(x.id, x.z.id); } catch (e) { /* resta da collegare a mano */ }
      caricaMeta(false);
    })();
  }, [meta, data, caricaMeta]);

  /* Ritorno dalla finestra di Facebook: ?meta=collegato oppure ?meta=errore&motivo=… */
  useEffect(() => {
    const q = new URLSearchParams(window.location.search);
    const esitoFb = q.get("meta");
    if (!esitoFb) return;
    if (esitoFb === "collegato") setMetaStato({ attesa: "", errore: "", esito: "Meta collegato. Ora abbina gli account ai clienti e premi Aggiorna ora." });
    else setMetaStato({ attesa: "", errore: `Collegamento a Meta non riuscito: ${q.get("motivo") || "riprova"}`, esito: "" });
    setCliAtt(""); setVista("impostazioni");
    window.history.replaceState(null, "", window.location.pathname);
  }, []);
  const collegaFacebook = async () => {
    if (!meta) return;
    setMetaStato({ attesa: "facebook", errore: "", esito: "" });
    try { await meta.collegaFacebook(); }
    catch (e) { setMetaStato({ attesa: "", errore: e.message || String(e), esito: "" }); }
  };
  const scollegaFacebook = async () => {
    if (!meta) return;
    try { await meta.scollegaFacebook(); await caricaMeta(false); setMetaStato({ attesa: "", errore: "", esito: "Meta scollegato: gli aggiornamenti si fermano finché non lo ricolleghi." }); }
    catch (e) { setMetaStato({ attesa: "", errore: e.message || String(e), esito: "" }); }
  };

  /* extra.account: aggiorna solo questi account (un cliente); chi: il cliente, per il pulsante della sua riga */
  const aggiornaMeta = async (azione, extra = {}, chi = "", solo = false) => {
    if (!meta || metaStato.attesa) return;
    setMetaStato({ attesa: azione === "account" ? "account" : "sync", errore: "", esito: "", cliente: chi, solo });
    try {
      const avvio = await meta.sincronizza(azione, extra);
      if (avvio && avvio.giaInCorso) setMetaStato((x) => ({ ...x, esito: "C'era già lo stesso aggiornamento in corso: aspetto che finisca." }));
      const mio = (avvio && avvio.sync) || null;
      if (azione === "account") {
        const info = await caricaMeta(false);
        setMetaStato({ attesa: "", errore: "", esito: `Elenco aggiornato: ${(info && info.account.length) || 0} account visibili.` });
        return;
      }
      /* La funzione lavora in sottofondo: aspetto che il registro dica che ha finito (al massimo 15 minuti),
         mostrando l'avanzamento. Se resta ferma più di un minuto (proseguimento automatico non partito),
         le chiedo io di continuare: funziona finché questa pagina è aperta. */
      for (let i = 0; i < 300; i++) {
        await new Promise((r) => setTimeout(r, 3000));
        const info = await meta.stato(mio);
        setMetaInfo(info);
        const u = info.seguita;
        if (!u || u.esito !== "in_corso") break;
        if (u.messaggio) setMetaStato((x) => ({ ...x, esito: u.messaggio }));
        const battito = u.dettagli && u.dettagli.battito;
        if (i % 10 === 9 && battito && Date.now() - new Date(battito).getTime() > 75e3) {
          try { await meta.sincronizza("continua", { sync: u.id }); } catch { /* riprovo al prossimo controllo */ }
        }
      }
      const info = await caricaMeta(true, mio);
      const u = info && (info.seguita || info.ultimaSync);
      if (!u || u.esito === "in_corso") setMetaStato({ attesa: "", errore: "", esito: "L'aggiornamento continua in sottofondo: i numeri arrivano alla prossima apertura." });
      else setMetaStato({ attesa: "", errore: u.esito === "errore" ? u.messaggio : "", esito: u.esito !== "errore" ? u.messaggio : "" });
    } catch (e) {
      setMetaStato({ attesa: "", errore: e.message || String(e), esito: "" });
    }
  };
  const collegaMeta = async (accountId, clienteId) => {
    if (!meta || !accountId) return;
    try {
      await meta.collega(accountId, clienteId);
      setMetaScelta((x) => ({ ...x, [clienteId]: "" }));
      await caricaMeta(false);
      setMetaStato({ attesa: "", errore: "", esito: "Account collegato: premi Aggiorna ora per scaricare i suoi dati." });
    } catch (e) { setMetaStato({ attesa: "", errore: e.message || String(e), esito: "" }); }
  };
  const scollegaMeta = async (accountId) => {
    if (!meta) return;
    try { await meta.scollega(accountId); await caricaMeta(false); }
    catch (e) { setMetaStato({ attesa: "", errore: e.message || String(e), esito: "" }); }
  };

  const esportaBackup = () => {
    /* La chiave API personale non entra nel file: un backup si inoltra, si archivia, si perde */
    const pacco = { formato: "crm-backup", versione: 6, generato: new Date().toISOString(), agenzia: data.agenzia.nome, dati: { ...data, aiKey: "" } };
    scaricaFile(`backup-${data.agenzia.nome.toLowerCase().replace(/[^a-z0-9]+/g, "-")}-${oggi()}.json`,
      JSON.stringify(pacco, null, 1), "application/json");
    setData((d) => ({ ...d, backup: { ...d.backup, ultimoFile: new Date().toISOString() } }));
  };

  /* Elimina un cliente e tutto ciò che gli appartiene, dopo aver detto esattamente cosa sparisce */
  /* ---------- richiesta di aggiornamento dati ---------- */
  /* L'app non può chiamare Meta da sola: prepara però la richiesta completa,
     con account, periodo già coperto e cosa manca, così basta incollarla in chat. */
  const componiRichiesta = (z) => {
    const ks = m.campagne.filter((k) => k.clienteId === z.id && k.archiviata !== "si");
    const attive = ks.filter((k) => ["attiva", "test"].includes(k.stato));
    const fin = (data.periodi || []).filter((pz) => ks.some((k) => k.id === pz.campagnaId));
    const ultima = fin.map((pz) => pz.fine).filter(Boolean).sort().slice(-1)[0] || "";
    const dal = ultima ? sommaGiorni(ultima, 1) : "";
    const righe = [];

    righe.push(`Aggiorna i dati Meta di ${z.nome}.`);
    righe.push("");
    righe.push(`Account pubblicitario: ${z.accountAdv || "DA INDICARE — copialo da Gestione inserzioni, inizia per act_"}`);
    righe.push(dal
      ? `Periodo da scaricare: dal ${dal} a ieri. Nel CRM ho già i dati fino al ${ultima}.`
      : "Periodo da scaricare: tutto lo storico disponibile. Nel CRM non ho ancora finestre caricate.");
    righe.push("");

    if (ks.length) {
      righe.push(`Campagne già presenti qui dentro (${ks.length}), da aggiornare per nome esatto:`);
      ks.slice(0, 25).forEach((k) => righe.push(`· ${k.nome}${["attiva", "test"].includes(k.stato) ? "  [attiva]" : ""}`));
      if (ks.length > 25) righe.push(`· e altre ${ks.length - 25}`);
    } else {
      righe.push("Non ho ancora campagne per questo cliente: creale tutte.");
    }
    righe.push("");

    righe.push("Cosa mi serve, nel formato di importazione del CRM:");
    righe.push("1. Campagne — totali del periodo, con piattaforma, obiettivo, stato, data di partenza, data di fine per quelle chiuse e budget giornaliero preso dall'ad set.");
    righe.push("2. Ad set e annunci — con il livello e l'ad set di appartenenza, così si agganciano fra loro.");
    righe.push(attive.length
      ? `3. Finestre — solo per le campagne attive (${attive.map((k) => k.nome).join(", ")}), tagliate sui giorni reali: 3 di Test, 7 di Matura, poi settimane.`
      : "3. Finestre — nessuna campagna attiva al momento, quindi bastano i totali.");
    righe.push("");
    righe.push("Per le campagne chiuse voglio solo i totali, non il dettaglio settimanale.");
    return righe.join("\n");
  };

  /* ---------- archiviazione dei clienti ---------- */
  /* Archiviare sposta tutto in un deposito: niente resta nei conti correnti,
     niente va perduto. Riattivare è la stessa operazione al contrario. */
  const COLL_CLI = ["contatti", "trattative", "campagne", "creativi", "rilevazioni", "periodi",
    "giornate", "vendite", "chiusure", "funnel", "file", "task", "annotazioni", "esperimenti", "automazioni", "profili"];

  const contaDati = (d, ids) => COLL_CLI.reduce((s, c) => s + (d[c] || []).filter((x) => ids.includes(x.clienteId)).length, 0);

  const archiviaCliente = (z) => {
    const ids = gruppo(z.id);
    const sotto = ids.length - 1;
    chiedi({
      titolo: `Archiviare ${z.nome}?`,
      testo: `Esce dall'elenco dei clienti attivi, dai totali dell'agenzia, dal semaforo, dai grafici e dagli avvisi.${sotto ? `\n\nCon lui vengono archiviati anche ${sotto} sottoclienti.` : ""}\n\nTutto lo storico — campagne, fasi, funnel, ciclo di vendita, chiusure, task, file — viene conservato e potrai reinserirlo se il cliente torna.`,
      etichetta: "Archivia",
      azione: () => setData((d) => {
        const dati = {};
        COLL_CLI.forEach((c) => { dati[c] = (d[c] || []).filter((x) => ids.includes(x.clienteId)); });
        const voce = { id: uid(), clienteId: z.id, quando: oggi(), motivo: "archiviazione", ids, dati };
        const nd = { ...d, deposito: [...(d.deposito || []), voce] };
        COLL_CLI.forEach((c) => { nd[c] = (d[c] || []).filter((x) => !ids.includes(x.clienteId)); });
        nd.clienti = d.clienti.map((x) => (ids.includes(x.id) ? { ...x, archiviato: "si", archiviatoIl: oggi() } : x));
        return nd;
      }),
    });
  };

  /* Solo l'advertising: campagne, ad set, finestre e rilevazioni.
     Funnel, ciclo di vendita, chiusure e task del cliente restano dove sono. */
  const COLL_ADV = ["campagne", "creativi", "periodi", "rilevazioni"];

  const archiviaAdv = (z, ids, etichettaGruppo) => setData((d) => {
    const dati = {};
    let n = 0;
    COLL_ADV.forEach((c) => {
      dati[c] = (d[c] || []).filter((x) => ids.includes(x.clienteId));
      n += dati[c].length;
    });
    if (!n) return d;
    const voce = { id: uid(), clienteId: z.id, quando: oggi(), motivo: "advertising", ids,
      etichetta: etichettaGruppo, dati };
    const nd = { ...d, deposito: [...(d.deposito || []), voce] };
    COLL_ADV.forEach((c) => { nd[c] = (d[c] || []).filter((x) => !ids.includes(x.clienteId)); });
    return nd;
  });

  const chiediArchiviaAdv = (z) => {
    const solo = [z.id];
    const conFigli = gruppo(z.id);
    const quante = (ids) => data.campagne.filter((k) => ids.includes(k.clienteId)).length;
    const dettaglio = (ids) => {
      const ks = data.campagne.filter((k) => ids.includes(k.clienteId));
      const sp = ks.reduce((s2, k) => s2 + (k.spesa || 0), 0)
        + (data.periodi || []).filter((pz) => ks.some((k) => k.id === pz.campagnaId)).reduce((s2, pz) => s2 + (pz.spesa || 0), 0);
      return `${ks.length} campagne e ${eur0(sp)} di spesa storica`;
    };
    if (!quante(conFigli)) {
      chiedi({ titolo: "Niente da archiviare", testo: "Questo cliente non ha campagne attive nell'elenco.", etichetta: "Ho capito", azione: () => {} });
      return;
    }
    const testoBase = "Campagne, ad set, finestre e rilevazioni escono da elenco, totali, grafici e avvisi. Funnel, ciclo di vendita, chiusure e task restano intatti.\n\nLo storico va in deposito e potrai reinserirlo quando vuoi.";
    if (figli(z.id).length) {
      chiedi({
        titolo: "Archiviare l'advertising?",
        testo: `${testoBase}\n\nSolo ${z.nome}: ${dettaglio(solo)}.\nCon i sottoclienti: ${dettaglio(conFigli)}.`,
        etichetta: "Anche i sottoclienti",
        azione: () => archiviaAdv(z, conFigli, "gruppo"),
        etichetta2: "Solo questo account",
        azione2: () => archiviaAdv(z, solo, "singolo"),
      });
    } else {
      chiedi({
        titolo: `Archiviare l'advertising di ${z.nome}?`,
        testo: `${testoBase}\n\nRiguarda ${dettaglio(solo)}.`,
        etichetta: "Archivia",
        azione: () => archiviaAdv(z, solo, "singolo"),
      });
    }
  };

  const riportaVoce = (voceId) => setData((d) => {
    const voce = (d.deposito || []).find((v2) => v2.id === voceId);
    if (!voce) return d;
    const nd = { ...d, deposito: (d.deposito || []).filter((v2) => v2.id !== voceId) };
    Object.keys(voce.dati).forEach((c) => {
      const esistenti = (d[c] || []).map((x) => x.id);
      nd[c] = [...(d[c] || []), ...(voce.dati[c] || []).filter((x) => !esistenti.includes(x.id))];
    });
    return nd;
  });

  const chiediRiattivaAdv = (voce) => {
    const n = COLL_ADV.reduce((s2, c) => s2 + (voce.dati[c] || []).length, 0);
    chiedi({
      titolo: "Riattivare l'advertising archiviato?",
      testo: `In deposito dal ${dataItA(voce.quando)} ci sono ${n} voci: ${(voce.dati.campagne || []).length} campagne con ad set, finestre e rilevazioni.\n\nCon lo storico torna tutto nei conti e negli avvisi. Ripartendo da zero i dati restano qui, e crei campagne nuove senza che le vecchie influenzino medie e confronti.`,
      etichetta: "Con lo storico",
      azione: () => riportaVoce(voce.id),
      etichetta2: "Riparti da zero",
      azione2: () => {},
    });
  };

  const riportaStorico = (clienteId) => setData((d) => {
    const voci = (d.deposito || []).filter((v) => v.clienteId === clienteId && v.motivo !== "advertising");
    if (!voci.length) return d;
    const nd = { ...d, deposito: (d.deposito || []).filter((v) => !(v.clienteId === clienteId && v.motivo !== "advertising")) };
    COLL_CLI.forEach((c) => {
      const esistenti = (d[c] || []).map((x) => x.id);
      const torna = voci.flatMap((v) => (v.dati[c] || [])).filter((x) => !esistenti.includes(x.id));
      nd[c] = [...(d[c] || []), ...torna];
    });
    return nd;
  });

  const riattivaCliente = (z) => {
    const voci = data.deposito.filter((v) => v.clienteId === z.id);
    const quanti = voci.reduce((s, v) => s + COLL_CLI.reduce((x, c) => x + (v.dati[c] || []).length, 0), 0);
    const sveglia = (d) => ({
      ...d, clienti: d.clienti.map((x) => (gruppoCompleto(z.id).includes(x.id) ? { ...x, archiviato: "no", archiviatoIl: "" } : x)),
    });
    chiedi({
      titolo: `Riattivare ${z.nome}?`,
      testo: `In archivio ci sono ${quanti} voci del suo storico, dal ${voci.length ? dataItA(voci[0].quando) : "—"}.\n\nPuoi riportarlo dentro con tutto lo storico, oppure ripartire da zero: in quel caso i dati restano nel deposito e potrai reinserirli in qualsiasi momento dalla sua scheda.`,
      etichetta: "Con lo storico",
      azione: () => { setData(sveglia); riportaStorico(z.id); },
      etichetta2: "Riparti da zero",
      azione2: () => setData(sveglia),
    });
  };

  const eliminaCliente = (z) => {
    const ids = gruppoCompleto(z.id);
    const sotto = ids.length - 1;
    const dentro = [
      ["contatti", data.contatti.filter((x) => ids.includes(x.clienteId)).length, "contatti"],
      ["trattative", data.trattative.filter((x) => ids.includes(x.clienteId)).length, "trattative"],
      ["campagne", data.campagne.filter((x) => ids.includes(x.clienteId)).length, "campagne"],
      ["creativi", data.creativi.filter((x) => ids.includes(x.clienteId)).length, "ad set e annunci"],
      ["vendite", data.vendite.filter((x) => ids.includes(x.clienteId)).length, "periodi commerciali"],
      ["chiusure", data.chiusure.filter((x) => ids.includes(x.clienteId)).length, "chiusure"],
      ["funnel", data.funnel.filter((x) => ids.includes(x.clienteId)).length, "funnel"],
      ["task", data.task.filter((x) => ids.includes(x.clienteId)).length, "task"],
      ["annotazioni", data.annotazioni.filter((x) => ids.includes(x.clienteId)).length, "annotazioni"],
      ["esperimenti", data.esperimenti.filter((x) => ids.includes(x.clienteId)).length, "esperimenti"],
      ["file", data.file.filter((x) => ids.includes(x.clienteId)).length, "file"],
      ["rilevazioni", data.rilevazioni.filter((x) => ids.includes(x.clienteId)).length, "rilevazioni storiche"],
    ];
    const inDeposito = (data.deposito || []).filter((v2) => ids.includes(v2.clienteId))
      .reduce((s2, v2) => s2 + Object.values(v2.dati || {}).reduce((x, l) => x + l.length, 0), 0);
    const elenco = [sotto > 0 ? `· ${sotto} sottoclienti, con tutto quello che contengono` : null,
      inDeposito > 0 ? `· ${inDeposito} voci di storico in deposito` : null,
      ...dentro.filter(([, n]) => n > 0).map(([, n, l]) => `· ${n} ${l}`)].filter(Boolean).join("\n") || "· nessun dato collegato";
    chiedi({
      titolo: `Eliminare ${z.nome}?`,
      testo: `Sparisce il cliente e tutto quello che contiene:\n\n${elenco}\n\nL'operazione non si può annullare. Se non hai un backup recente, annulla e scaricalo prima.`,
      etichetta: "Elimina definitivamente", pericolo: true,
      azione: () => {
        setData((d) => {
          const nd = { ...d, clienti: d.clienti.filter((x) => !ids.includes(x.id)) };
          dentro.forEach(([coll]) => { nd[coll] = d[coll].filter((x) => !ids.includes(x.clienteId)); });
          nd.automazioni = d.automazioni.filter((a) => !ids.includes(a.clienteId));
          nd.profili = (d.profili || []).filter((p) => !ids.includes(p.clienteId));
          /* Via anche l'eventuale storico in deposito: eliminare deve eliminare davvero */
          nd.deposito = (d.deposito || []).filter((v2) => !ids.includes(v2.clienteId));
          nd.giornate = (d.giornate || []).filter((g) => !ids.includes(g.clienteId));
          nd.periodi = (d.periodi || []).filter((pz) => !ids.includes(pz.clienteId));
          nd.eventi = (d.eventi || []).filter((ev) => !ids.includes(ev.clienteId));
          return nd;
        });
        /* Se stavi guardando un altro account resti dove sei */
        if (ids.includes(cliAtt)) { setCliAtt(""); setVista("clienti"); }
      },
    });
  };

  const m = useMemo(() => {
    if (!data) return null;
    const tutteCamp = campagneCalcolate(data);
    /* Le archiviate restano consultabili ma non entrano in nessun calcolo */
    const campagne = tutteCamp.filter((k) => k.archiviata !== "si");
    const archiviate = tutteCamp.filter((k) => k.archiviata === "si");
    const attive = campagne.filter((k) => ["attiva", "test"].includes(k.stato));
    const tutti = campagne.flatMap((k) => veri(k.al).map((a) => ({ ...a, campagna: k.nome, clienteId: k.clienteId, campagnaId: k.id })));

    /* I task collegano calendario e avvisi: scaduti o urgenti in scadenza
       non devono vivere solo in una lista che puoi non aprire. */
    (data.task || []).filter((t) => t.stato !== "fatto" && t.scadenza).forEach((t) => {
      const giorni = giorniTra(oggi(), t.scadenza) - 1;
      const urgente = ["urgente", "alta"].includes(t.priorita);
      if (t.scadenza < oggi()) {
        tutti.push({
          gravita: urgente ? "critica" : "attenzione", clienteId: t.clienteId, taskId: t.id,
          metrica: "Task", valore: `${Math.abs(giorni)} ${Math.abs(giorni) === 1 ? "giorno" : "giorni"}`, soglia: `era il ${dataIt(t.scadenza)}`,
          titolo: `Task in ritardo · ${t.titolo}`,
          cosa: `Priorità ${(PRIORITA.find((p) => p.v === t.priorita) || {}).l || "media"}, scaduto il ${dataItA(t.scadenza)}.`,
          causa: "Un task scaduto che resta aperto di solito significa che è più grande di come l'hai scritto, o che non è davvero tuo.",
          azione: "Aprilo dal calendario o dalla sezione Task: chiudilo, spostalo con una data vera, o spezzalo in due.",
        });
      } else if (urgente && giorni <= 1) {
        tutti.push({
          gravita: "nota", clienteId: t.clienteId, taskId: t.id,
          metrica: "Task", valore: giorni === 0 ? "oggi" : "domani", soglia: dataIt(t.scadenza),
          titolo: `In scadenza · ${t.titolo}`,
          cosa: `Task a priorità ${(PRIORITA.find((p) => p.v === t.priorita) || {}).l.toLowerCase()} con scadenza ${giorni === 0 ? "oggi" : "domani"}.`,
          causa: "Compare qui perché l'hai segnato come urgente o ad alta priorità.",
          azione: "Se non lo fai oggi, spostagli la data: una scadenza che slitta in silenzio smette di significare qualcosa.",
        });
      }
    });
    const portafoglio = data.clienti.filter((z) => z.stato !== "chiuso" && z.archiviato !== "si").map((z) => conValuta(valutaCliente(data, z.id), () => {
      const ks = attive.filter((k) => k.clienteId === z.id);
      const sp = ks.reduce((s, k) => s + k.d.sp, 0), rc = ks.reduce((s, k) => s + k.d.rc, 0), cv = ks.reduce((s, k) => s + k.d.cv, 0);
      const ld = ks.reduce((s, k) => s + k.d.ld, 0);
      const tR = ks.filter((k) => k.bmk.roasMin > 0), tC = ks.filter((k) => tettoCpl(k.bmk) > 0);
      const oR = tR.length ? tR.reduce((s, k) => s + k.bmk.roasMin, 0) / tR.length : 0;
      const oC = tC.length ? tC.reduce((s, k) => s + tettoCpl(k.bmk), 0) / tC.length : 0;
      let modo = "nd", perf = 0, testo = "dati insufficienti", obiettivo = "";
      if (rc > 0 && oR > 0) { modo = "roas"; perf = (rc / sp) / oR; testo = `ROAS ${fmt(rc / sp)}×`; obiettivo = `obiettivo ${fmt(oR)}×`; }
      /* Dove ci sono lead il confronto si fa sul CPL: è il costo che il cliente vede per primo */
      else if (ld > 0 && oC > 0) { modo = "cpa"; perf = oC / (sp / ld); testo = `CPL ${eur(sp / ld)}`; obiettivo = `tetto ${eur(oC)}`; }
      else if (cv > 0 && oC > 0) { modo = "cpa"; perf = oC / (sp / cv); testo = `CPA ${eur(sp / cv)}`; obiettivo = `tetto ${eur(oC)}`; }
      else if (sp > 0) { testo = `${eur(sp)} spesi`; obiettivo = cv ? `${int(cv)} conversioni` : "nessuna conversione"; }
      return { ...z, sp, rc, cv, modo, perf, testo, obiettivo, nCamp: ks.length, nAvvisi: ks.reduce((s, k) => s + veri(k.al).length, 0) };
    }));
    return {
      campagne, archiviate, attive, portafoglio, avvisi: tutti,
      spesa: attive.reduce((s, k) => s + k.d.sp, 0),
      ricavi: attive.reduce((s, k) => s + k.d.rc, 0),
      conv: attive.reduce((s, k) => s + k.d.cv, 0),
      critici: tutti.filter((a) => a.gravita === "critica").length,
    };
  }, [data]);

  if (!data || !m) return <div className="crm" data-tema={tema}><style>{CSS}</style><div className="empty" style={{ padding: 60 }}>Carico i tuoi dati…</div></div>;

  const cliente = (id) => data.clienti.find((z) => z.id === id);

  /* ---------- Meta Ads: pannelli (scheda cliente e Impostazioni) ---------- */
  const metaAccount = (metaInfo && metaInfo.account) || [];
  const metaColl = (metaInfo && metaInfo.collegamenti) || [];
  const metaFb = (metaInfo && metaInfo.facebook) || null;
  /* Giorni rimasti arrotondati per eccesso; scaduto (-1) solo quando la data è davvero passata */
  const metaFbGiorni = metaFb && metaFb.scadenza
    ? (new Date(metaFb.scadenza) < new Date() ? -1 : Math.ceil((new Date(metaFb.scadenza) - Date.now()) / 864e5)) : null;
  const metaDiChi = new Map(metaColl.map((c) => [String(c.account_id), c.cliente_id]));
  const metaQuando = (s2) => (s2 ? new Date(s2).toLocaleString(LOCALE, { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }) : "mai");
  const metaStatoAcc = (a) => ACCOUNT_STATI[Number(a.stato)] || (a.stato == null ? "—" : `stato ${a.stato}`);
  const metaEtichetta = (a) => `${a.nome || a.account_id} · act_${a.account_id}${a.valuta ? ` · ${a.valuta}` : ""}${accountUtilizzabile(a) ? "" : ` · ${metaStatoAcc(a)}`}`;
  const metaSuggeriti = (z) => suggerisciAccount(z, metaAccount.filter((a) => !metaDiChi.has(String(a.account_id)) && accountUtilizzabile(a))).slice(0, 5);
  /* Gli account da aggiornare per un cliente: i suoi, e con i sottoclienti anche i loro */
  const metaIdDi = (id, conSotto) => {
    const ids = conSotto ? [id, ...discendenti(id)] : [id];
    return metaColl.filter((c) => ids.includes(c.cliente_id)).map((c) => String(c.account_id));
  };
  const metaAggiornaCliente = (z, conSotto) => {
    const ids = metaIdDi(z.id, conSotto);
    if (ids.length) aggiornaMeta(undefined, { account: ids }, z.id, !conSotto && discendenti(z.id).length > 0);
  };
  /* Regole per account (quali campagne, se entra nell'aggiornamento di tutti): stanno nel collegamento */
  const metaRegolaDi = (accId) => metaColl.find((c) => String(c.account_id) === String(accId)) || {};
  const metaRegolaBreve = (accId) => {
    const r = metaRegolaDi(accId), n = (r.campagne_scelte || []).length;
    return [r.campagne_modo === "attive" ? "solo attive" : r.campagne_modo === "scelte" ? `${n} ${n === 1 ? "campagna scelta" : "campagne scelte"}` : "", r.automatico === false ? "solo a mano" : ""].filter(Boolean).join(" · ");
  };
  const META_STATO_CAMP = { ACTIVE: "attiva", PAUSED: "in pausa", CAMPAIGN_PAUSED: "in pausa", ARCHIVED: "archiviata", DELETED: "eliminata", IN_PROCESS: "in revisione", WITH_ISSUES: "con problemi" };
  const salvaRegolaMeta = async (accId, campi, avviso = "Impostazione salvata: vale dal prossimo aggiornamento.") => {
    try {
      await meta.regola(accId, campi);
      await caricaMeta(false);
      setMetaStato({ attesa: "", errore: "", esito: avviso });
      return true;
    } catch (e) { setMetaStato({ attesa: "", errore: e.message || String(e), esito: "" }); return false; }
  };
  const apriSceltaCampagne = async (accId) => {
    setMetaRegola({ account: accId, lista: null, scelte: (metaRegolaDi(accId).campagne_scelte || []).map(String), filtro: "", errore: "" });
    try {
      const lista = await meta.campagne(accId);
      setMetaRegola((r) => (r && r.account === accId ? { ...r, lista } : r));
    } catch (e) { setMetaRegola((r) => (r && r.account === accId ? { ...r, errore: e.message || String(e) } : r)); }
  };
  const salvaSceltaCampagne = async () => {
    const r = metaRegola;
    if (!r || !r.scelte.length) return;
    const ok = await salvaRegolaMeta(r.account, { campagne_modo: "scelte", campagne_scelte: r.scelte },
      `Salvato: per questo account il CRM scaricherà solo ${r.scelte.length === 1 ? "la campagna scelta" : `le ${r.scelte.length} campagne scelte`}. Premi Aggiorna questo cliente per caricarle subito.`);
    if (ok) setMetaRegola(null);
  };
  /* Riga sotto l'account nella scheda cliente: campagne da scaricare e aggiornamento automatico */
  const metaRegolaRiga = (a) => {
    const r = metaRegolaDi(a.account_id), modo = r.campagne_modo || "recenti";
    const aperta = metaRegola && metaRegola.account === a.account_id ? metaRegola : null;
    const n = (r.campagne_scelte || []).length;
    return (
      <>
        <div className="meta-regola">
          <span className="sub">Campagne da scaricare</span>
          <select value={aperta ? "scelte" : modo} aria-label={`Campagne da scaricare per ${a.nome || a.account_id}`}
            onChange={(e) => {
              const v = e.target.value;
              if (v === "scelte") apriSceltaCampagne(a.account_id);
              else { setMetaRegola(null); salvaRegolaMeta(a.account_id, { campagne_modo: v }); }
            }}>
            <option value="recenti">Attive o con spesa negli ultimi 90 giorni</option>
            <option value="attive">Solo le attive</option>
            <option value="scelte">Solo quelle che scelgo{modo === "scelte" ? ` (${n})` : ""}</option>
          </select>
          {modo === "scelte" && !aperta && <button className="btn ghost mini" onClick={() => apriSceltaCampagne(a.account_id)}>Cambia la scelta</button>}
          <label title="Se lo togli, l'account si aggiorna solo con il pulsante di questo cliente">
            <input type="checkbox" checked={r.automatico !== false}
              onChange={(e) => salvaRegolaMeta(a.account_id, { automatico: e.target.checked }, e.target.checked
                ? "L'account torna nell'aggiornamento delle 9 e in Aggiorna ora."
                : "L'account resta fuori dall'aggiornamento delle 9 e da Aggiorna ora: si aggiorna solo con il pulsante di questo cliente.")} />
            Nell'aggiornamento automatico (alle 9 e con Aggiorna ora)
          </label>
        </div>
        {aperta && (() => {
          const filtro = aperta.filtro.trim().toLowerCase();
          const lista = (aperta.lista || []).filter((c) => !filtro || (c.nome || "").toLowerCase().includes(filtro) || c.id.includes(filtro));
          const cambia = (id, si) => setMetaRegola((x) => ({ ...x, scelte: si ? [...new Set([...x.scelte, id])] : x.scelte.filter((y) => y !== id) }));
          return (
            <div style={{ marginTop: 10 }}>
              {aperta.errore ? <div className="meta-msg" data-tipo="errore">{aperta.errore}</div>
                : !aperta.lista ? <div className="sub">Leggo le campagne da Meta…</div>
                : !aperta.lista.length ? <div className="sub">Meta non restituisce campagne per questo account.</div>
                : (
                  <>
                    <div className="meta-riga">
                      <input type="search" placeholder="Cerca per nome" value={aperta.filtro} style={{ maxWidth: 260 }}
                        onChange={(e) => setMetaRegola((x) => ({ ...x, filtro: e.target.value }))} aria-label="Cerca campagna" />
                      <button className="btn ghost mini" onClick={() => setMetaRegola((x) => ({ ...x, scelte: [...new Set([...x.scelte, ...x.lista.filter((c) => c.stato === "ACTIVE").map((c) => c.id)])] }))}>Aggiungi le attive</button>
                      <button className="btn ghost mini" onClick={() => setMetaRegola((x) => ({ ...x, scelte: [] }))}>Togli tutte</button>
                    </div>
                    <div className="meta-camp">
                      {lista.map((c) => (
                        <label key={c.id}>
                          <input type="checkbox" checked={aperta.scelte.includes(c.id)} onChange={(e) => cambia(c.id, e.target.checked)} />
                          <span className="nome"><b>{c.nome || c.id}</b><span className="sub"> · {META_STATO_CAMP[c.stato] || (c.stato || "").toLowerCase()}{c.inizio ? ` · dal ${dataItA(c.inizio)}` : ""}</span></span>
                          <span className="sub mono">{c.spesa90 ? `${Number(c.spesa90).toLocaleString(LOCALE, { minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: true })} ${a.valuta || ""} in 90 gg` : "nessuna spesa in 90 gg"}</span>
                        </label>
                      ))}
                      {!lista.length && <div className="sub" style={{ padding: 10 }}>Nessuna campagna con questo nome.</div>}
                    </div>
                  </>
                )}
              <div className="meta-riga">
                <button className="btn ghost mini" disabled={!aperta.scelte.length} onClick={salvaSceltaCampagne}>Salva la scelta ({aperta.scelte.length})</button>
                <button className="btn ghost mini" onClick={() => setMetaRegola(null)}>Annulla</button>
                <span className="sub">Le campagne non scelte che sono già nel CRM restano: se non ti servono eliminale, non verranno ricaricate.</span>
              </div>
            </div>
          );
        })()}
      </>
    );
  };
  const metaDelCliente = (id) => metaColl.filter((c) => c.cliente_id === id)
    .map((c) => metaAccount.find((a) => String(a.account_id) === String(c.account_id)) || { account_id: String(c.account_id), nome: "", stato: null });

  const metaMessaggi = () => (
    <>
      {metaStato.errore && <div className="meta-msg" data-tipo="errore">{metaStato.errore}</div>}
      {metaStato.esito && <div className="meta-msg" data-tipo="ok">{metaStato.esito}</div>}
      {metaStato.attesa === "sync" && <div className="meta-msg">Sto scaricando i dati da Meta: può volerci qualche minuto se gli account sono tanti.</div>}
    </>
  );
  const metaPulsanti = () => (
    <>
      <button className="btn ghost mini" disabled={!!metaStato.attesa || !metaFb} onClick={() => aggiornaMeta("account")}>
        {metaStato.attesa === "account" ? "Leggo gli account…" : "Aggiorna l'elenco degli account"}</button>
      <button className="btn ghost mini" disabled={!!metaStato.attesa || !metaColl.length || !metaFb} onClick={() => aggiornaMeta()}
        title={metaColl.length ? "Scarica adesso da Meta i dati fino a ieri" : "Collega prima almeno un account"}>
        {metaStato.attesa === "sync" ? "Aggiornamento in corso…" : "Aggiorna ora"}</button>
    </>
  );
  /* Menu per collegare un account: prima i suggeriti per nome, poi tutti; quelli già di un altro cliente non si scelgono */
  const metaCollega = (z) => {
    const sugg = metaSuggeriti(z);
    /* Nessuna scelta già fatta: i suggeriti stanno in cima al menu, ma l'account lo sceglie l'utente */
    const scelto = metaScelta[z.id] || "";
    const tutti = [...metaAccount].sort((a, b) => (a.nome || "").localeCompare(b.nome || ""));
    return (
      <div className="meta-riga">
        <select value={scelto} onChange={(e) => setMetaScelta((x) => ({ ...x, [z.id]: e.target.value }))} aria-label={`Account Meta per ${z.nome}`}>
          <option value="">— scegli un account —</option>
          {sugg.length > 0 && (
            <optgroup label="Suggeriti per nome">
              {sugg.map((a) => <option key={"s" + a.account_id} value={a.account_id}>{metaEtichetta(a)}</option>)}
            </optgroup>
          )}
          <optgroup label="Tutti gli account">
            {tutti.map((a) => {
              const altro = metaDiChi.get(String(a.account_id));
              return <option key={a.account_id} value={a.account_id} disabled={!!altro}>{metaEtichetta(a)}{altro ? ` — collegato a ${(cliente(altro) || {}).nome || "un altro cliente"}` : ""}</option>;
            })}
          </optgroup>
        </select>
        <button className="btn ghost mini" disabled={!scelto} onClick={() => collegaMeta(scelto, z.id)}>Collega</button>
      </div>
    );
  };
  const metaChip = (a) => (
    <span className="meta-chip" key={a.account_id} title={`act_${a.account_id}${a.ultimo_errore ? ` · ultimo errore: ${a.ultimo_errore}` : ""}`}>
      {a.nome || `act_${a.account_id}`}{a.valuta ? <span className="sub">{a.valuta}</span> : null}
      {metaRegolaBreve(a.account_id) ? <span className="sub" title="Si cambia dalla scheda del cliente">· {metaRegolaBreve(a.account_id)}</span> : null}
      {a.ultimo_errore ? <span className="neg" aria-label="errore">!</span> : null}
      <button onClick={() => scollegaMeta(a.account_id)} title="Scollega" aria-label={`Scollega ${a.nome || a.account_id}`}>×</button>
    </span>
  );

  const PannelloMetaCliente = (z) => {
    if (!meta) return (
      <section className="card">
        <header><h2>Meta Ads</h2></header>
        <div className="empty" style={{ padding: 18 }}>Il collegamento con Meta funziona nella versione online del CRM.</div>
      </section>
    );
    const miei = metaDelCliente(z.id);
    return (
      <section className="card">
        <header>
          <h2>Meta Ads</h2>
          <span className="sp eyebrow">ultimo aggiornamento {metaQuando(metaInfo && metaInfo.ultimaSync && metaInfo.ultimaSync.fine)}</span>
          {/* Con sottoclienti due pulsanti: solo il cliente (i suoi account) o anche i sottoclienti */}
          {metaFb && discendenti(z.id).length > 0 && metaIdDi(z.id, false).length > 0 && (
            <button className="btn ghost mini" disabled={!!metaStato.attesa} onClick={() => metaAggiornaCliente(z, false)}
              title="Scarica adesso solo gli account di questo cliente, senza i sottoclienti">
              {metaStato.attesa === "sync" && metaStato.cliente === z.id && metaStato.solo ? "Aggiornamento in corso…" : "Aggiorna solo questo cliente"}
            </button>
          )}
          {metaFb && metaIdDi(z.id, true).length > 0 && (
            <button className="btn ghost mini" disabled={!!metaStato.attesa} onClick={() => metaAggiornaCliente(z, true)}
              title={discendenti(z.id).length ? "Scarica adesso i dati di questo cliente e dei suoi sottoclienti" : "Scarica adesso i dati di questo cliente"}>
              {metaStato.attesa === "sync" && metaStato.cliente === z.id && !metaStato.solo ? "Aggiornamento in corso…" : discendenti(z.id).length ? "Aggiorna cliente e sottoclienti" : "Aggiorna questo cliente"}
            </button>
          )}
          {metaPulsanti()}
        </header>
        {metaMessaggi()}
        {!metaAccount.length ? (
          <div className="empty" style={{ padding: 18 }}>
            <b>Nessun account letto da Meta</b>
            {metaFb ? `Premi "Aggiorna l'elenco degli account": il CRM legge gli account a cui hai accesso e ti propone quelli che somigliano a ${z.nome}.` : "Prima collega il tuo profilo Facebook da Impostazioni → Collegamento con Meta Ads."}
          </div>
        ) : (
          <table className="tbl"><tbody>
            {miei.map((a) => (
              <tr key={a.account_id}>
                <td><b>{a.nome || `act_${a.account_id}`}</b><div className="sub mono">act_{a.account_id}{a.business ? ` · ${a.business}` : ""}</div></td>
                <td className="sub">{a.valuta || "—"} · {metaStatoAcc(a)}</td>
                <td className="sub">{a.ultimo_errore ? <span className="neg">{a.ultimo_errore}</span> : `dati del ${metaQuando(a.ultimo_sync)}`}</td>
                <td className="r"><button className="btn danger mini" onClick={() => scollegaMeta(a.account_id)}>Scollega</button></td>
              </tr>
            )).flatMap((riga, i) => [riga, <tr key={"r" + miei[i].account_id}><td colSpan={4} style={{ borderTop: 0, paddingTop: 0 }}>{metaRegolaRiga(miei[i])}</td></tr>])}
            <tr>
              <td colSpan={4}>
                <div className="sub" style={{ marginBottom: 6 }}>{miei.length ? "Collega un altro account a questo cliente" : "Collega l'account pubblicitario di questo cliente"}</div>
                {metaCollega(z)}
              </td>
            </tr>
          </tbody></table>
        )}
      </section>
    );
  };

  const PannelloTelegram = () => {
    if (!telegram) return null;
    const ore = Array.from({ length: 28 }, (_, i) => { const m = 7 * 60 + i * 30; return `${String(Math.floor(m / 60)).padStart(2, "0")}:${String(m % 60).padStart(2, "0")}`; });
    return (
      <section className="card">
        <header>
          <h2>Bot Telegram</h2>
          <span className="sp eyebrow">{tgChat.length ? `collegato · ${tgChat.length} ${tgChat.length === 1 ? "chat" : "chat"}` : "non collegato"}</span>
          <button className="btn ghost mini" disabled={!!tgStato.attesa} onClick={() => azioneTelegram(() => telegram.collega(), (r) => ({ link: r.link, esito: r.anthropic ? "" : "Manca ANTHROPIC_API_KEY tra i segreti di Supabase: il bot risponde ai pulsanti ma non capisce i messaggi." }))}>
            {tgChat.length ? "Collega un'altra chat" : "Collega Telegram"}</button>
          {tgChat.length > 0 && <button className="btn ghost mini" disabled={!!tgStato.attesa} onClick={() => azioneTelegram(() => telegram.prova(), (r) => ({ esito: `Messaggio di prova inviato a ${r.inviati} chat.` }))}>Manda una prova</button>}
        </header>
        <div className="meta-msg sub">
          Scrivi al bot cosa fare e lui crea task, note nel diario e spostamenti; rispondi alle domande sui dati.
          Ogni mattina manda il riepilogo dei task con i pulsanti Fatto e Domani, e avvisa prima dei task con un orario.
          {tgStato.applicati > 0 ? ` In questa sessione ha già inserito ${tgStato.applicati} ${tgStato.applicati === 1 ? "elemento" : "elementi"}.` : ""}
        </div>
        {tgStato.errore && <div className="meta-msg" data-tipo="errore">{tgStato.errore}</div>}
        {tgStato.esito && <div className="meta-msg" data-tipo="ok">{tgStato.esito}</div>}
        {tgStato.link && (
          <div className="meta-msg" style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
            <span>Apri il link dal telefono (o da Telegram sul computer) e premi <b>Avvia</b>: vale 30 minuti.</span>
            <a className="btn ghost mini" href={tgStato.link} target="_blank" rel="noopener noreferrer">Apri Telegram</a>
            <span className="sub mono" style={{ userSelect: "all" }}>{tgStato.link}</span>
            <button className="btn ghost mini" onClick={() => { setTgStato((x) => ({ ...x, link: "" })); caricaTelegram(); }}>Fatto, ho premuto Avvia</button>
          </div>
        )}
        {tgChat.length > 0 && (
          <div className="wrap"><table className="tbl">
            <thead><tr><th>Chat</th><th>Riepilogo del mattino</th><th>Sabato e domenica</th><th>Avviso prima dei task con orario</th><th className="r" /></tr></thead>
            <tbody>
              {tgChat.map((c) => (
                <tr key={c.chat_id}>
                  <td><b>{c.utente || "Telegram"}</b><div className="sub">{c.attivo ? "attiva" : "in pausa"} · dal {dataIt(String(c.collegato || "").slice(0, 10))}</div></td>
                  <td>
                    <select value={c.ora_riepilogo} aria-label="Ora del riepilogo" style={{ width: "auto" }}
                      onChange={(e) => azioneTelegram(() => telegram.preferenze(c.chat_id, { ora_riepilogo: e.target.value }), "Ora del riepilogo salvata.")}>
                      {[...new Set([...ore, c.ora_riepilogo])].sort().map((o) => <option key={o} value={o}>{o}</option>)}
                    </select>
                  </td>
                  <td>
                    <label className="riga" style={{ gap: 6 }}><input type="checkbox" style={{ width: 16, height: 16, accentColor: "var(--acc)" }} checked={!!c.weekend}
                      onChange={(e) => azioneTelegram(() => telegram.preferenze(c.chat_id, { weekend: e.target.checked }), "Salvato.")} /> anche nel weekend</label>
                  </td>
                  <td>
                    <select value={c.preavviso} aria-label="Preavviso" style={{ width: "auto" }}
                      onChange={(e) => azioneTelegram(() => telegram.preferenze(c.chat_id, { preavviso: Number(e.target.value) }), "Preavviso salvato.")}>
                      {[0, 10, 15, 30, 60, 120].map((v) => <option key={v} value={v}>{v ? `${v} minuti prima` : "all'ora esatta"}</option>)}
                    </select>
                  </td>
                  <td className="r" style={{ whiteSpace: "nowrap" }}>
                    <button className="btn ghost mini" onClick={() => azioneTelegram(() => telegram.preferenze(c.chat_id, { attivo: !c.attivo }), c.attivo ? "Chat in pausa: niente messaggi finché non la riattivi." : "Chat riattivata.")}>{c.attivo ? "Metti in pausa" : "Riattiva"}</button>
                    <button className="btn danger mini" onClick={() => azioneTelegram(() => telegram.scollega(c.chat_id), "Chat scollegata.")}>Scollega</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table></div>
        )}
      </section>
    );
  };

  /* Software esterni che mandano task al CRM (il primo: Boschetto, di Kesia) */
  const cambiaInteg = (id, fn) => setData((d) => ({ ...d, integrazioni: (d.integrazioni || []).map((x) => (x.id === id ? fn(x) : x)) }));
  /* Abbinamento "nome del cliente nel software esterno → sottocliente del CRM":
     sposta anche le task già arrivate che non hai spostato a mano */
  const abbinaInteg = (integ, chiave, clienteId) => setData((d) => {
    const vecchio = ((integ.abbinamenti || {})[chiave]) || integ.clienteBase || "";
    const abbinamenti = { ...(integ.abbinamenti || {}) };
    if (clienteId) abbinamenti[chiave] = clienteId; else delete abbinamenti[chiave];
    const nuovo = clienteId || integ.clienteBase || "";
    return {
      ...d,
      integrazioni: (d.integrazioni || []).map((x) => (x.id === integ.id ? { ...x, abbinamenti } : x)),
      task: d.task.map((t) => (t.esterno && t.esterno.integrazioneId === integ.id && nomeChiave(t.esterno.cliente) === chiave && (t.clienteId || "") === vecchio
        ? { ...t, clienteId: nuovo } : t)),
    };
  });
  const baseInteg = (integ, clienteId) => setData((d) => {
    const vecchio = integ.clienteBase || "";
    const ab = integ.abbinamenti || {};
    return {
      ...d,
      integrazioni: (d.integrazioni || []).map((x) => (x.id === integ.id ? { ...x, clienteBase: clienteId } : x)),
      task: d.task.map((t) => (t.esterno && t.esterno.integrazioneId === integ.id && !ab[nomeChiave(t.esterno.cliente)] && (t.clienteId || "") === vecchio
        ? { ...t, clienteId } : t)),
    };
  });
  const istruzioniIntegrazione = (nome, indirizzo, chiave) => `Collegamento ${nome} → CRM NC Media

Ogni volta che in ${nome} una task assegnata a me viene creata, modificata, completata o eliminata, manda una richiesta HTTP:

POST ${indirizzo}
Authorization: Bearer ${chiave || "<LA CHIAVE CHE TI HO MANDATO A PARTE>"}
Content-Type: application/json

{
  "evento": "task.creata",          // oppure "task.aggiornata" o "task.eliminata"
  "task": {
    "id": "id della task in ${nome} (sempre lo stesso per la stessa task)",
    "titolo": "Testo della task",
    "descrizione": "Dettagli (facoltativo)",
    "cliente": "Nome del cliente o progetto in ${nome}",
    "scadenza": "2026-10-23",       // AAAA-MM-GG, facoltativa; va bene anche 2026-10-23T15:30
    "ora": "15:30",                 // facoltativa
    "priorita": "alta",             // urgente | alta | media | bassa
    "stato": "da fare",             // da fare | in corso | in attesa | fatto
    "url": "https://… link alla task in ${nome} (facoltativo)"
  }
}

Per mandarne più di una insieme: { "tasks": [ {…}, {…} ] }. Al primo collegamento conviene mandare tutte le task aperte assegnate a me.
Mandare di nuovo la stessa task (stesso id) la aggiorna, non la duplica. Risposta: { "ok": true, "ricevute": N }.

Ritorno (i cambi che faccio io nel CRM: stato, scadenza, ora, priorità, titolo). Due modi, si può usare uno o l'altro:
1) ${nome} chiede periodicamente (per esempio ogni 5 minuti):
   POST ${indirizzo}   con la stessa chiave
   { "azione": "modifiche", "dal": <ultimo numero ricevuto, 0 la prima volta> }
   Risposta: { "ok": true, "modifiche": [ { "id": 12, "esternoId": "id della task", "titolo", "scadenza", "ora", "priorita", "stato", "cambiati": ["stato"], "quando" } ], "ultimo": 12 }
   Gli stati arrivano come: dafare | incorso | attesa | fatto. Salva "ultimo" e usalo come "dal" la volta dopo.
2) Oppure ${nome} mi dà un suo indirizzo https (e se vuole una chiave): il CRM ci manda subito
   POST { "evento": "task.modificata_nel_crm", "modifiche": [ … come sopra … ] }   con Authorization: Bearer <la vostra chiave>.

La chiave è riservata: va tenuta sul server di ${nome} (non nel codice visibile dal browser).`;
  const PannelloIntegrazioni = () => {
    if (!integrazioni) return null;
    const locali = data.integrazioni || [];
    const copia = (testo, esito) => { try { navigator.clipboard.writeText(testo); setIntStato((x) => ({ ...x, esito })); } catch { /* selezione manuale */ } };
    const crea = async () => {
      const nome = (intStato.nome || "").trim() || "Boschetto";
      const r = await azioneIntegrazione(() => integrazioni.crea(nome), "");
      if (!r) return;
      setData((d) => ({ ...d, integrazioni: [...(d.integrazioni || []).filter((x) => x.id !== r.id), { id: r.id, nome, clienteBase: intStato.base || "", abbinamenti: {}, nomiVisti: {} }] }));
      setIntStato((x) => ({ ...x, chiave: { id: r.id, chiave: r.chiave }, indirizzo: r.indirizzo || x.indirizzo }));
    };
    const opzioniCli = (base) => (base ? ordinaGerarchia().filter(({ z }) => gruppo(base).includes(z.id)) : ordinaGerarchia());
    const selCli = (valore, onChange, vuoto, base = "", etichetta = "") => (
      <select value={valore || ""} aria-label={etichetta} onChange={(e) => onChange(e.target.value)} style={{ width: "auto", maxWidth: 280 }}>
        <option value="">{vuoto}</option>
        {opzioniCli(base).map(({ z, liv }) => <option key={z.id} value={z.id}>{liv ? `${"\u00A0".repeat(liv * 4)}↳ ` : ""}{z.nome}</option>)}
      </select>
    );
    return (
      <section className="card">
        <header>
          <h2>Software esterni</h2>
          <span className="sp eyebrow">{(() => { const n = intLista.filter((x) => x.tipo !== "mcp").length; return n ? `${n} ${n === 1 ? "collegato" : "collegati"}` : "nessuno"; })()}</span>
        </header>
        <div className="meta-msg sub">
          Un altro software (per esempio Boschetto di Kesia) manda qui le task che ti assegna: entrano nei Task sotto il sottocliente giusto,
          e quando le cambi nel CRM (fatto, scadenza, priorità) il cambio torna indietro.
        </div>
        {intStato.errore && <div className="meta-msg" data-tipo="errore">{intStato.errore}</div>}
        {intStato.esito && <div className="meta-msg" data-tipo="ok">{intStato.esito}</div>}
        {intStato.chiave && (() => {
          const integ = intLista.find((x) => x.id === intStato.chiave.id) || { nome: "" };
          return (
            <div className="meta-msg" style={{ display: "grid", gap: 8 }}>
              <div><b>Chiave di {integ.nome}</b>: si vede solo adesso. Copiala e mandala a chi gestisce il software (meglio in un messaggio separato dalle istruzioni).</div>
              <div className="sub mono" style={{ userSelect: "all", wordBreak: "break-all" }}>{intStato.chiave.chiave}</div>
              <div className="riga" style={{ gap: 8, flexWrap: "wrap" }}>
                <button className="btn ghost mini" onClick={() => copia(intStato.chiave.chiave, "Chiave copiata.")}>Copia la chiave</button>
                <button className="btn ghost mini" onClick={() => copia(istruzioniIntegrazione(integ.nome, intStato.indirizzo, intStato.chiave.chiave), "Istruzioni con la chiave copiate.")}>Copia istruzioni con la chiave</button>
                <button className="btn ghost mini" onClick={() => setIntStato((x) => ({ ...x, chiave: null }))}>Fatto, l'ho salvata</button>
              </div>
            </div>
          );
        })()}
        {intLista.filter((x) => x.tipo !== "mcp").map((x) => {
          const loc = locali.find((l) => l.id === x.id) || { id: x.id, nome: x.nome, clienteBase: "", abbinamenti: {}, nomiVisti: {} };
          const nomi = Object.entries(loc.nomiVisti || {}).sort((a, b) => a[1].localeCompare(b[1]));
          const quante = data.task.filter((t) => t.esterno && t.esterno.integrazioneId === x.id && !t.esterno.eliminata);
          const aperte = quante.filter((t) => t.stato !== "fatto").length;
          const base = loc.clienteBase || "";
          return (
            <div key={x.id} style={{ borderTop: "1px solid var(--line)", padding: "12px 18px", display: "grid", gap: 10 }}>
              <div className="riga" style={{ gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                <b>{x.nome}</b>
                <span className="pill">{x.attiva ? "attivo" : "in pausa"}</span>
                <span className="sub">chiave …{x.chiave_fine} · {x.ultimo_contatto ? `ultimo contatto ${dataIt(String(x.ultimo_contatto).slice(0, 10))} ${String(x.ultimo_contatto).slice(11, 16)}` : "non ha ancora mandato niente"} · {quante.length} task ({aperte} aperte)</span>
                <span className="sp" />
                <button className="btn ghost mini" disabled={!!intStato.attesa} onClick={() => copia(istruzioniIntegrazione(x.nome, intStato.indirizzo, ""), "Istruzioni copiate (senza chiave).")}>Copia istruzioni</button>
                <button className="btn ghost mini" disabled={!!intStato.attesa} onClick={async () => { if (!window.confirm(`Nuova chiave per ${x.nome}? Quella di adesso smette subito di funzionare.`)) return; const r = await azioneIntegrazione(() => integrazioni.rigenera(x.id), ""); if (r) setIntStato((s2) => ({ ...s2, chiave: { id: x.id, chiave: r.chiave } })); }}>Nuova chiave</button>
                <button className="btn ghost mini" disabled={!!intStato.attesa} onClick={() => azioneIntegrazione(() => integrazioni.attiva(x.id, !x.attiva), x.attiva ? "In pausa: le task in arrivo vengono rifiutate." : "Riattivato.")}>{x.attiva ? "Metti in pausa" : "Riattiva"}</button>
                <button className="btn danger mini" disabled={!!intStato.attesa} onClick={async () => { if (!window.confirm(`Scollegare ${x.nome}? Le task già arrivate restano nel CRM.`)) return; const r = await azioneIntegrazione(() => integrazioni.elimina(x.id), "Scollegato."); if (r) setData((d) => ({ ...d, integrazioni: (d.integrazioni || []).filter((l) => l.id !== x.id) })); }}>Scollega</button>
              </div>
              <label className="riga" style={{ gap: 8, flexWrap: "wrap" }}>
                <span className="sub">Cliente principale (dove vanno le task senza abbinamento)</span>
                {selCli(base, (v) => baseInteg(loc, v), "— nessuno (task dell'agenzia) —", "", "Cliente principale")}
              </label>
              <div>
                <div className="sub" style={{ marginBottom: 6 }}>Clienti di {x.nome} → sottoclienti del CRM {nomi.length ? "" : "· i nomi compaiono qui quando arriva la prima task"}</div>
                {nomi.length > 0 && (
                  <div className="wrap"><table className="tbl">
                    <thead><tr><th>In {x.nome}</th><th>Nel CRM</th><th className="r">Task</th></tr></thead>
                    <tbody>
                      {nomi.map(([k, originale]) => (
                        <tr key={k}>
                          <td>{originale}</td>
                          <td>{selCli((loc.abbinamenti || {})[k], (v) => abbinaInteg(loc, k, v), base ? `${nomeCli(base)} (principale)` : "— cliente principale —", base, `Sottocliente per ${originale}`)}</td>
                          <td className="r">{quante.filter((t) => nomeChiave(t.esterno.cliente) === k).length}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table></div>
                )}
              </div>
              <details>
                <summary className="sub" style={{ cursor: "pointer" }}>Ritorno: {x.ritorno ? `il CRM avvisa ${x.url_ritorno}` : `${x.nome} legge i cambi quando vuole (nessun indirizzo di ritorno)`}</summary>
                <form className="riga" style={{ gap: 8, flexWrap: "wrap", marginTop: 8 }} onSubmit={(e) => { e.preventDefault(); const f = new FormData(e.currentTarget);
                  const chiave = String(f.get("chiave") || "");
                  azioneIntegrazione(() => integrazioni.ritorno(x.id, String(f.get("url") || ""), chiave ? chiave : undefined), "Indirizzo di ritorno salvato."); }}>
                  <input name="url" defaultValue={x.url_ritorno || ""} placeholder="https://… indirizzo di Boschetto (facoltativo)" style={{ minWidth: 280, flex: 1 }} />
                  <input name="chiave" type="password" autoComplete="off" placeholder="chiave che vuole ricevere (facoltativa)" style={{ width: 240 }} />
                  <button className="btn ghost mini" disabled={!!intStato.attesa}>Salva</button>
                </form>
              </details>
            </div>
          );
        })}
        <div className="riga" style={{ gap: 8, flexWrap: "wrap", padding: "12px 18px", borderTop: "1px solid var(--line)" }}>
          <input value={intStato.nome} onChange={(e) => setIntStato((x) => ({ ...x, nome: e.target.value }))} placeholder="Nome del software" aria-label="Nome del software" style={{ width: 180 }} />
          {selCli(intStato.base, (v) => setIntStato((x) => ({ ...x, base: v })), "— cliente principale —", "", "Cliente principale")}
          <button className="btn" disabled={!!intStato.attesa} onClick={crea}>Collega un software</button>
        </div>
      </section>
    );
  };

  /* Connettore del CRM per Claude (funzione crm-mcp): il Claude di un cliente legge il CRM e ci salva le analisi */
  const istruzioniProgetto = (nome) => `Nel CRM NC Media c'è il connettore «CRM NC Media». Quando analizzi le winner ads di ${nome || "questo cliente"}:
1. leggi_benchmark con cliente "${nome || "nome del cliente"}" e usa quei riferimenti (costo per lead, per appuntamento e per cliente chiuso, minimo di clienti);
2. prendi da Hyros il report di attribuzione per annuncio del periodo (ultimi 30 giorni, salvo richiesta diversa) e chiama salva_dati_ads con TUTTI gli annunci con spesa, attivi e in pausa: id, nome, creativa (il nome o il numero comune a tutte le copie della stessa creatività, es. "71"), campagna, adset, pubblico ("freddo" o "retargeting", sempre indicato), attiva, spesa, lead, appuntamenti, clienti, incassato, e "periodo" con le date. Il CRM somma le copie della stessa creatività e fa una classifica per il freddo e una per il retargeting;
3. scrivi la tua analisi (winner ads, cosa scalare, cosa spegnere, cosa rifare) e salvala con salva_analisi, tipo "winner_ads";
4. se serve un'azione concreta, crea un task con crea_task.
Puoi leggere anche i dati del CRM con dati_cliente (campagne, ciclo di vendita, task) e le analisi precedenti con leggi_analisi.`;
  const PannelloConnettore = () => {
    if (!integrazioni) return null;
    const miei = intLista.filter((x) => x.tipo === "mcp");
    /* Il nome che il Claude di Boutique del Fitness deve usare (se il cliente c'è) */
    const nomeBdf = (attivi.find((z) => /boutique/i.test(z.nome)) || {}).nome || "";
    const crea = async () => {
      const r = await azioneIntegrazione(() => integrazioni.crea("Claude", "mcp"), "");
      if (r) setIntStato((x) => ({ ...x, mcp: { id: r.id, url: r.connettore } }));
    };
    const copia = (testo, esito) => { try { navigator.clipboard.writeText(testo); setIntStato((x) => ({ ...x, esito })); } catch { /* selezione manuale */ } };
    return (
      <section className="card">
        <header>
          <h2>Connettore per Claude</h2>
          <span className="sp eyebrow">{miei.length ? (miei.some((x) => x.ultimo_contatto) ? "in uso" : "creato, mai usato") : "non creato"}</span>
          {!miei.length && <button className="btn" disabled={!!intStato.attesa} onClick={crea}>Crea il connettore</button>}
        </header>
        <div className="meta-msg sub">
          Il CRM diventa uno strumento dei tuoi Claude su claude.ai: il Claude di un cliente (per esempio quello di Boutique del Fitness, collegato a Hyros)
          legge i dati del CRM e salva qui i numeri degli annunci e le sue analisi, che trovi nella sezione Winner ads del cliente. Può anche creare task.
        </div>
        {intStato.errore && <div className="meta-msg" data-tipo="errore">{intStato.errore}</div>}
        {intStato.esito && <div className="meta-msg" data-tipo="ok">{intStato.esito}</div>}
        {intStato.mcp && (
          <div className="meta-msg" style={{ display: "grid", gap: 8 }}>
            {(() => {
              /* La chiave va nell'intestazione Authorization (claude.ai la conserva cifrata), l'indirizzo resta senza chiave */
              const base = intStato.mcp.url.replace(/\/crmnc_[^/]*$/, ""), chiave = (intStato.mcp.url.match(/crmnc_[^/]*$/) || [""])[0];
              return (
                <>
                  <div>In claude.ai: <b>Impostazioni → Connettori → Aggiungi connettore personalizzato</b>, nome «CRM NC Media».</div>
                  <div><b>1. Indirizzo del server</b> (senza chiave):</div>
                  <div className="sub mono" style={{ userSelect: "all", wordBreak: "break-all" }}>{base}</div>
                  <div><b>2. Autenticazione</b>: «Nessun accesso». <b>3. Intestazioni della richiesta → Aggiungi intestazione</b>: nome <span className="mono">Authorization</span>, valore:</div>
                  <div className="sub mono" style={{ userSelect: "all", wordBreak: "break-all" }}>Bearer {chiave}</div>
                  <div className="sub">Il valore si vede solo adesso: è la chiave, non mandarla a nessuno. Poi salva e, nella chat o nel Progetto del cliente, attiva il connettore dal menu degli strumenti.</div>
                  <div className="riga" style={{ gap: 8, flexWrap: "wrap" }}>
                    <button className="btn ghost mini" onClick={() => copia(base, "Indirizzo copiato.")}>Copia l'indirizzo</button>
                    <button className="btn ghost mini" onClick={() => copia("Bearer " + chiave, "Valore dell'intestazione copiato.")}>Copia il valore dell'intestazione</button>
                  </div>
                </>
              );
            })()}
            <div className="riga" style={{ gap: 8, flexWrap: "wrap" }}>
              <button className="btn ghost mini" onClick={() => copia(istruzioniProgetto(nomeBdf), "Istruzioni per il Progetto copiate.")}>Copia istruzioni per il Progetto</button>
              <button className="btn ghost mini" onClick={() => setIntStato((x) => ({ ...x, mcp: null }))}>Fatto, l'ho aggiunto</button>
            </div>
          </div>
        )}
        {miei.map((x) => (
          <div key={x.id} className="riga" style={{ gap: 10, flexWrap: "wrap", alignItems: "center", padding: "12px 18px", borderTop: "1px solid var(--line)" }}>
            <b>CRM NC Media</b>
            <span className="pill">{x.attiva ? "attivo" : "in pausa"}</span>
            <span className="sub">chiave …{x.chiave_fine} · {x.ultimo_contatto ? `ultimo uso ${dataIt(String(x.ultimo_contatto).slice(0, 10))} ${String(x.ultimo_contatto).slice(11, 16)}` : "mai usato"}</span>
            <span className="sp" />
            <button className="btn ghost mini" onClick={() => copia(istruzioniProgetto(nomeBdf), "Istruzioni per il Progetto copiate.")}>Copia istruzioni per il Progetto</button>
            <button className="btn ghost mini" disabled={!!intStato.attesa} onClick={async () => { if (!window.confirm("Nuovo indirizzo per il connettore? Quello di adesso smette subito di funzionare e va sostituito in claude.ai.")) return; const r = await azioneIntegrazione(() => integrazioni.rigenera(x.id), ""); if (r) setIntStato((st) => ({ ...st, mcp: { id: x.id, url: r.connettore } })); }}>Nuovo indirizzo</button>
            <button className="btn ghost mini" disabled={!!intStato.attesa} onClick={() => azioneIntegrazione(() => integrazioni.attiva(x.id, !x.attiva), x.attiva ? "Connettore in pausa: Claude riceve un rifiuto finché non lo riattivi." : "Connettore riattivato.")}>{x.attiva ? "Metti in pausa" : "Riattiva"}</button>
            <button className="btn danger mini" disabled={!!intStato.attesa} onClick={() => { if (window.confirm("Eliminare il connettore? Le analisi già salvate restano.")) azioneIntegrazione(() => integrazioni.elimina(x.id), "Connettore eliminato: toglilo anche da claude.ai."); }}>Elimina</button>
          </div>
        ))}
      </section>
    );
  };

  /* Piattaforme pubblicitarie collegabili al CRM: oggi Meta, le altre predisposte */
  const PIATTAFORME_COLLEGABILI = [
    { id: "meta", nome: "Meta", sotto: "Facebook e Instagram Ads", sigla: "M", colore: "#1877F2" },
    { id: "google", nome: "Google Ads", sotto: "Search, Display, YouTube", sigla: "G", colore: "#EA4335" },
    { id: "tiktok", nome: "TikTok Ads", sotto: "TikTok for Business", sigla: "T", colore: "#25F4EE" },
    { id: "linkedin", nome: "LinkedIn Ads", sotto: "Campaign Manager", sigla: "in", colore: "#0A66C2" },
  ];
  /* Fusi orari: i più usati in cima, poi tutti quelli che il browser conosce */
  const FUSI_COMUNI = ["Europe/Rome", "Europe/London", "Europe/Madrid", "Europe/Paris", "Europe/Berlin", "America/New_York", "America/Chicago", "America/Denver", "America/Phoenix", "America/Los_Angeles", "America/Anchorage", "Pacific/Honolulu", "America/Toronto", "America/Sao_Paulo", "America/Mexico_City", "Asia/Dubai", "Australia/Sydney"];
  const etichettaFuso = (f) => {
    try {
      const ora = new Intl.DateTimeFormat("it-IT", { timeZone: f, hour: "2-digit", minute: "2-digit" }).format(new Date());
      return `${f.replace(/_/g, " ")} · ora ${ora}`;
    } catch { return f; }
  };
  const PannelloOrari = () => {
    if (!spazio) return null;
    let tutti = [];
    try { tutti = Intl.supportedValuesOf("timeZone"); } catch { tutti = []; }
    const altri = tutti.filter((f) => !FUSI_COMUNI.includes(f));
    const ore = Array.from({ length: 24 }, (_, h) => h);
    return (
      <section className="card">
        <header>
          <h2>Fuso orario</h2>
          <span className="sp eyebrow">vale per tutto lo spazio di lavoro</span>
        </header>
        {orari.errore && <div className="meta-msg" data-tipo="errore">{orari.errore}</div>}
        {orari.esito && <div className="meta-msg" data-tipo="ok">{orari.esito}</div>}
        <div className="riga" style={{ gap: 14, flexWrap: "wrap", padding: "14px 18px", alignItems: "flex-end" }}>
          <label style={{ display: "grid", gap: 4 }}>
            <span className="eyebrow">Fuso orario dell'agenzia</span>
            <select value={orari.fuso} onChange={(e) => salvaOrari({ fuso: e.target.value })} style={{ width: "auto", maxWidth: 360 }}>
              <optgroup label="Più usati">{[...new Set([orari.fuso, ...FUSI_COMUNI])].map((f) => <option key={f} value={f}>{etichettaFuso(f)}</option>)}</optgroup>
              {altri.length > 0 && <optgroup label="Tutti">{altri.map((f) => <option key={f} value={f}>{etichettaFuso(f)}</option>)}</optgroup>}
            </select>
          </label>
          <label style={{ display: "grid", gap: 4 }}>
            <span className="eyebrow">Aggiornamento di Meta</span>
            <select value={orari.ora_meta} onChange={(e) => salvaOrari({ ora_meta: Number(e.target.value) })} style={{ width: "auto" }}>
              {ore.map((h) => <option key={h} value={h}>ogni giorno alle {String(h).padStart(2, "0")}:00</option>)}
            </select>
          </label>
          {fusoDelBrowser() !== orari.fuso && (
            <button className="btn ghost mini" onClick={() => salvaOrari({ fuso: fusoDelBrowser() })}>Usa il fuso di questo computer ({fusoDelBrowser().replace(/_/g, " ")})</button>
          )}
        </div>
        <div className="meta-msg sub">
          All'ora scelta, nel fuso dell'agenzia, il CRM scarica da Meta i dati fino a ieri. Lo stesso fuso vale per il riepilogo e i promemoria del bot Telegram
          e per la data delle analisi salvate dal connettore per Claude. Per uno spazio di lavoro nuovo il CRM propone il fuso del computer di chi lo crea.
        </div>
      </section>
    );
  };

  const PannelloAspetto = () => (
    <section className="card">
      <header><h2>Aspetto</h2><span className="sp eyebrow">vale per questo browser</span></header>
      <div className="riga" style={{ gap: 10, padding: "14px 18px", flexWrap: "wrap" }}>
        {[["scuro", Moon, "Modalità scura"], ["chiaro", Sun, "Modalità chiara"]].map(([v, I, l]) => (
          <button key={v} className={tema === v ? "btn" : "btn ghost"} aria-pressed={tema === v} onClick={() => cambiaTema(v)}
            style={{ display: "inline-flex", alignItems: "center", gap: 8 }}>
            <I size={16} strokeWidth={1.9} aria-hidden="true" />{l}
          </button>
        ))}
      </div>
    </section>
  );
  const PannelloPiattaforme = () => {
    if (!meta) return null;
    const stato = (p) => {
      if (p.id !== "meta") return { testo: "in arrivo", tipo: "" };
      if (!metaInfo) return { testo: "verifico…", tipo: "" };
      if (!metaFb) return { testo: "non collegato", tipo: "" };
      if (metaFbGiorni !== null && metaFbGiorni < 0) return { testo: "accesso scaduto", tipo: "neg" };
      return { testo: `collegato come ${metaFb.utente_meta || "—"}`, tipo: metaFbGiorni !== null && metaFbGiorni < 7 ? "neg" : "pos" };
    };
    return (
      <section className="card">
        <header>
          <h2>Piattaforme pubblicitarie</h2>
          <span className="sp eyebrow">collega i profili da cui il CRM legge campagne e risultati, in sola lettura</span>
        </header>
        {metaStato.attesa === "facebook" && <div className="meta-msg">Apro la pagina di accesso di Meta…</div>}
        {metaFb && metaFbGiorni !== null && metaFbGiorni < 7 && (
          <div className="meta-msg" data-tipo="errore">
            {metaFbGiorni < 0 ? "L'accesso a Meta è scaduto: gli aggiornamenti sono fermi. Premi Rinnova il collegamento." : `L'accesso a Meta scade fra ${metaFbGiorni} ${metaFbGiorni === 1 ? "giorno" : "giorni"}: premi Rinnova il collegamento, bastano due clic.`}
          </div>
        )}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(240px, 1fr))", gap: 12, padding: "14px 18px" }}>
          {PIATTAFORME_COLLEGABILI.map((p) => {
            const st = stato(p), attiva = p.id === "meta";
            return (
              <div key={p.id} style={{ border: "1px solid var(--line)", borderRadius: 10, padding: 14, display: "grid", gap: 10, opacity: attiva ? 1 : 0.6 }}>
                <div className="riga" style={{ gap: 10 }}>
                  <span aria-hidden="true" style={{ width: 34, height: 34, borderRadius: 8, background: p.colore, color: p.id === "tiktok" ? "#111" : "#fff", display: "inline-flex", alignItems: "center", justifyContent: "center", fontWeight: 700, flex: "none" }}>{p.sigla}</span>
                  <div><b>{p.nome}</b><div className="sub">{p.sotto}</div></div>
                </div>
                <div className={"sub " + st.tipo}>{st.testo}
                  {attiva && metaFb && metaFb.scadenza && metaFbGiorni !== null && metaFbGiorni >= 0 ? ` · scade il ${dataItA(metaFb.scadenza.slice(0, 10))}` : ""}</div>
                {attiva ? (
                  <div className="riga" style={{ gap: 8, flexWrap: "wrap" }}>
                    {!metaFb ? (
                      <button className="btn" disabled={!!metaStato.attesa} onClick={collegaFacebook}>{metaStato.attesa === "facebook" ? "Apro Meta…" : "Collega a Meta"}</button>
                    ) : (
                      <>
                        <button className="btn ghost mini" disabled={!!metaStato.attesa} onClick={collegaFacebook}>{metaFbGiorni !== null && metaFbGiorni < 7 ? "Rinnova il collegamento" : "Ricollega"}</button>
                        <button className="btn danger mini" onClick={() => { if (window.confirm("Scollegare il profilo Meta? Gli aggiornamenti delle campagne si fermano finché non lo ricolleghi; i dati già scaricati restano.")) scollegaFacebook(); }}>Scollega</button>
                      </>
                    )}
                  </div>
                ) : (
                  <button className="btn ghost mini" disabled title="Collegamento in preparazione">In arrivo</button>
                )}
              </div>
            );
          })}
        </div>
        <div className="meta-msg sub">
          Con «Collega a Meta» accedi con il tuo profilo Facebook: il CRM vede, in sola lettura, gli account pubblicitari che gestisci. Poi, nel riquadro qui sotto, abbini ogni account al suo cliente.
          Il collegamento vale per questo spazio di lavoro e dura 60 giorni; una settimana prima della scadenza il CRM ti avvisa.
        </div>
      </section>
    );
  };

  const PannelloMetaAgenzia = () => {
    if (!meta) return null;
    const u = metaInfo && metaInfo.ultimaSync;
    const righe = ordinaGerarchia().filter(({ z }) => z.archiviato !== "si" && z.stato !== "chiuso");
    const collegati = righe.filter(({ z }) => metaColl.some((c) => c.cliente_id === z.id)).length;
    const conErrori = metaAccount.filter((a) => metaDiChi.has(String(a.account_id)) && a.ultimo_errore);
    return (
      <section className="card">
        <header>
          <h2>Collegamento con Meta Ads</h2>
          <span className="sp eyebrow">{collegati} di {righe.length} clienti collegati · {metaAccount.length} account visibili · ▸ apre i sottoclienti</span>
          {metaPulsanti()}
        </header>
        {!metaFb && metaInfo && <div className="meta-msg">Il profilo Meta non è collegato: premi «Collega a Meta» nel riquadro Piattaforme pubblicitarie qui sopra.</div>}
        <div className="meta-msg sub">
          Ogni giorno alle {String(orari.ora_meta).padStart(2, "0")}:00{orari.fuso !== "Europe/Rome" ? ` (${orari.fuso.replace(/_/g, " ")})` : ""} il CRM scarica da Meta i dati fino a ieri per tutti gli account collegati, tranne quelli segnati «solo a mano» nella scheda del cliente.
          {" "}Ultimo aggiornamento: <b>{u ? `${metaQuando(u.fine || u.avvio)} · ${u.esito === "in_corso" ? "in corso" : u.esito}` : "mai"}</b>{u && u.messaggio ? ` — ${u.messaggio}` : ""}
        </div>
        {metaMessaggi()}
        {conErrori.length > 0 && (
          <div className="meta-msg" data-tipo="errore">
            {conErrori.map((a) => <div key={a.account_id}><b>{a.nome}</b> ({(cliente(metaDiChi.get(String(a.account_id))) || {}).nome || "—"}): {a.ultimo_errore}</div>)}
          </div>
        )}
        {!metaAccount.length ? (
          <div className="empty" style={{ padding: 18 }}>
            <b>Nessun account letto da Meta</b>
            {metaFb ? "Premi \"Aggiorna l'elenco degli account\". Se resta vuoto, il profilo collegato non ha accesso ad account pubblicitari." : "Premi \"Collega a Meta\" qui sopra: appena collegato, il CRM legge gli account e ti propone i collegamenti con i clienti."}
          </div>
        ) : (
          <div className="wrap"><table className="tbl">
            <thead><tr><th>Cliente</th><th>Account collegati</th><th>Collega</th><th className="r">Dati</th></tr></thead>
            <tbody>
              {alberoCli("meta", righe).map(({ z, liv, nSotto }) => {
                const miei = metaDelCliente(z.id);
                const sotto = nSotto ? gruppo(z.id).slice(1) : [];
                const sottoColl = sotto.filter((id) => metaColl.some((c) => c.cliente_id === id)).length;
                return (
                  <tr key={z.id}>
                    <td style={{ paddingLeft: 14 + liv * 18 }}>
                      {frecciaCli("meta", z.id, nSotto)}
                      {liv ? <span className="sub">↳ </span> : null}<b>{z.nome}</b>
                      {nSotto > 0 && <div className="sub">con {int(sotto.length)} {sotto.length === 1 ? "sottocliente" : "sottoclienti"} · {sottoColl ? `${int(sottoColl)} collegati` : "nessuno collegato"}</div>}
                    </td>
                    <td>{miei.length ? miei.map(metaChip) : <span className="sub">nessuno</span>}</td>
                    <td style={{ minWidth: 260 }}>{metaCollega(z)}</td>
                    <td className="r" style={{ whiteSpace: "nowrap" }}>
                      {miei.length > 0 && (
                        <>
                          <button className="btn ghost mini" disabled={!!metaStato.attesa || !metaFb} onClick={() => metaAggiornaCliente(z, false)}
                            title="Scarica adesso solo gli account di questo cliente">
                            {metaStato.attesa === "sync" && metaStato.cliente === z.id ? "Aggiorno…" : "Aggiorna"}
                          </button>
                          <div className="sub" style={{ marginTop: 4 }}>
                            {(() => {
                              const date = miei.map((a) => a.ultimo_sync).filter(Boolean).sort();
                              return date.length < miei.length ? "mai aggiornato" : `dati del ${metaQuando(date[0])}`;
                            })()}
                          </div>
                        </>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table></div>
        )}
      </section>
    );
  };
  /* Valute negli elenchi con più clienti.
     valutaComune: la valuta condivisa da tutti quei clienti, o quella attiva se sono diverse
     (in quel caso i totali restano nella valuta dell'agenzia, come spiega la Guida).
     mapValuta: come .map(), ma ogni riga è scritta nella valuta del suo cliente. */
  const valutaComune = (ids) => {
    const v = [...new Set(ids.filter(Boolean).map((id) => valutaCliente(data, id)))];
    return v.length === 1 ? v[0] : VALUTA;
  };
  const mapCon = (lista, valutaDi, fn) => lista.map((x, i) => conValuta(valutaDi(x), () => fn(x, i)));
  const mapValuta = (lista, cliDi, fn) => mapCon(lista, (x) => valutaCliente(data, cliDi(x)), fn);
  const nomeCli = (id) => (cliente(id) || {}).nome || "Agenzia";
  const nomeArea = (id) => (data.aree.find((a) => a.id === id) || {}).nome || "senza area";
  const attivo = cliAtt ? cliente(cliAtt) : null;
  /* Data di partenza: il primo giorno di attività pubblicitaria registrata.
     Se non ci sono campagne, resta la data in cui hai creato la scheda. */
  const partenza = (id) => {
    const date = data.campagne.filter((k) => k.clienteId === id && k.archiviata !== "si" && k.inizio).map((k) => k.inizio);
    const z = data.clienti.find((x) => x.id === id);
    const prima = date.length ? date.sort()[0] : "";
    return { data: prima || (z ? z.creato : ""), daCampagne: !!prima };
  };

  const figli = (id) => data.clienti
    .filter((z) => z.padreId === id && z.archiviato !== "si")
    .sort(perCognome);
  const discendenti = (id) => figli(id).flatMap((z) => [z.id, ...discendenti(z.id)]);
  const gruppo = (id) => [id, ...discendenti(id)];
  /* Come gruppo, ma include anche i sottoclienti archiviati: serve a riattivare ed eliminare
     senza lasciare figli dimenticati in archivio o dati orfani. */
  const discendentiTutti = (id) => data.clienti.filter((z) => z.padreId === id).flatMap((z) => [z.id, ...discendentiTutti(z.id)]);
  const gruppoCompleto = (id) => [id, ...discendentiTutti(id)];
  const haFigli = attivo ? figli(attivo.id).length > 0 : false;
  const ambito = attivo ? (consolidato && haFigli ? gruppo(cliAtt) : [cliAtt]) : [];
  /* Elenco piatto ma ordinato: ogni padre seguito dai suoi sottoclienti */
  const attivi = data.clienti.filter((z) => z.archiviato !== "si");
  const ordinaGerarchia = (padre = "", liv = 0) =>
    attivi.filter((z) => (z.padreId || "") === padre)
      .sort((a, b) => (padre ? perCognome(a, b) : 0))   /* i sottoclienti per cognome, i principali nell'ordine tuo */
      .flatMap((z) => [{ z, liv }, ...ordinaGerarchia(z.id, liv + 1)]);
  /* Menù a tendina dei sottoclienti, come in Advertising: nelle tabelle con più clienti si vedono
     i clienti principali e la freccia apre i loro sottoclienti. sez distingue le sezioni. */
  const apertoCli = (sez, id) => !!apertiCli[sez + ":" + id];
  const frecciaCli = (sez, id, n) => (n > 0 ? (
    <button type="button" className="adv-freccia" aria-expanded={apertoCli(sez, id)} aria-label={apertoCli(sez, id) ? "Chiudi i sottoclienti" : "Apri i sottoclienti"}
      title={apertoCli(sez, id) ? "Chiudi i sottoclienti" : "Apri i sottoclienti"}
      onClick={(e) => { e.stopPropagation(); setApertiCli((x) => ({ ...x, [sez + ":" + id]: !x[sez + ":" + id] })); }}>{apertoCli(sez, id) ? "▾" : "▸"}</button>
  ) : null);
  /* Da un elenco gerarchico ({ z, liv } in ordine) alle righe visibili, ognuna con quanti sottoclienti ha */
  const alberoCli = (sez, righe) => {
    const out = [];
    let chiuso = Infinity;
    righe.forEach((r, i) => {
      if (r.liv > chiuso) return;
      let n = 0;
      for (let j = i + 1; j < righe.length && righe[j].liv > r.liv; j++) n++;
      chiuso = apertoCli(sez, r.z.id) ? Infinity : r.liv;
      out.push({ ...r, nSotto: n });
    });
    return out;
  };
  /* Tabelle di totali: una riga per cliente principale con i numeri di tutto il gruppo; aperta,
     sotto compaiono il solo cliente principale e ogni sottocliente. calc(ids) → numeri o null. */
  const righeGruppi = (sez, calc, ordina = null) => {
    const gruppi = ordinaGerarchia().filter((x) => x.liv === 0).map(({ z }) => {
      const sotto = ordinaGerarchia(z.id, 1);
      const tot = calc([z.id, ...sotto.map((x) => x.z.id)]);
      if (!tot) return null;
      const r = { z, liv: 0, key: z.id, nSotto: sotto.length, ...tot };
      if (!sotto.length || !apertoCli(sez, z.id)) return [r];
      const propria = calc([z.id]);
      return [r, ...(propria ? [{ z, liv: 1, key: z.id + ":solo", solo: true, ...propria }] : []),
        ...sotto.map((x) => { const d = calc([x.z.id]); return d ? { z: x.z, liv: x.liv, key: x.z.id, ...d } : null; }).filter(Boolean)];
    }).filter(Boolean);
    if (ordina) gruppi.sort((a, b) => ordina(a[0], b[0]));
    return gruppi.flat();
  };
  const nomeGruppo = (sez, r) => (
    <>
      {r.liv === 0 && frecciaCli(sez, r.z.id, r.nSotto)}
      {r.liv ? <span className="sub">↳ </span> : null}<b>{r.z.nome}</b>
      {r.solo && <div className="sub">solo il cliente principale</div>}
      {r.liv === 0 && r.nSotto > 0 && <div className="sub">con {int(r.nSotto)} {r.nSotto === 1 ? "sottocliente" : "sottoclienti"}</div>}
    </>
  );
  const logoAgenzia = data.agenzia.logo || LOGO_AG;

  const S = {
    contatti: data.contatti.filter((c) => ambito.includes(c.clienteId)),
    lead: data.contatti.filter((c) => ambito.includes(c.clienteId) && c.tipo === "lead"),
    trattative: data.trattative.filter((t) => ambito.includes(t.clienteId)),
    campagne: m.campagne.filter((k) => ambito.includes(k.clienteId)),
    task: data.task.filter((t) => ambito.includes(t.clienteId)),
    vendite: data.vendite.filter((v) => ambito.includes(v.clienteId)),
    chiusure: data.chiusure.filter((c) => ambito.includes(c.clienteId)),
    creativi: data.creativi.filter((c) => ambito.includes(c.clienteId)),
    annotazioni: data.annotazioni.filter((a) => ambito.includes(a.clienteId)),
    esperimenti: data.esperimenti.filter((e) => ambito.includes(e.clienteId)),
    funnel: data.funnel.filter((f) => ambito.includes(f.clienteId)),
    file: data.file.filter((f) => ambito.includes(f.clienteId)),
  };
  S.avvisi = S.campagne.flatMap((k) => veri(k.al).map((a) => ({ ...a, campagna: k.nome, clienteId: k.clienteId, campagnaId: k.id })));
  const contatto = (id) => data.contatti.find((c) => c.id === id);

  const kpiCli = (() => {
    const at = S.campagne.filter((k) => ["attiva", "test"].includes(k.stato));
    const sp = at.reduce((s, k) => s + k.d.sp, 0), rc = at.reduce((s, k) => s + k.d.rc, 0), cv = at.reduce((s, k) => s + k.d.cv, 0);
    const aperte = S.trattative.filter((t) => !["vinto", "perso"].includes(t.fase));
    return {
      sp, rc, cv, roas: sp ? rc / sp : 0, cpa: cv ? sp / cv : 0,
      pipeline: aperte.reduce((s, t) => s + t.valore, 0),
      previsione: aperte.reduce((s, t) => s + t.valore * faseDi(t.fase).p, 0),
      funnel: [
        { k: "Lead", n: S.lead.length, sub: "TOTALI" },
        { k: "Qualificati", n: S.lead.filter((c) => ["qualificato", "opportunita", "cliente"].includes(c.stato)).length, sub: "MQL" },
        { k: "Opportunità", n: S.trattative.length, sub: "TRATTATIVE" },
        { k: "Clienti", n: S.trattative.filter((t) => t.fase === "vinto").length, sub: "VINTE" },
      ],
    };
  })();

  /* I clienti archiviati non si propongono più, salvo quello già scritto nel record che stai modificando */
  const optCliTutti = [{ v: "", l: "— interna all'agenzia —" }, ...data.clienti
    .filter((z) => z.archiviato !== "si" || (ed && ed.draft && ed.draft.clienteId === z.id))
    .map((z) => ({ v: z.id, l: z.archiviato === "si" ? `${z.nome} (archiviato)` : z.nome }))];  const optAree = data.aree.map((a) => ({ v: a.id, l: a.nome }));
  const optCampCli = [{ v: "", l: "— nessuna —" }, ...S.campagne.map((k) => ({ v: k.id, l: k.nome }))];
  /* Gli ad set selezionabili: quelli della campagna scelta, o tutti quelli del cliente */
  const optAdset = [{ v: "", l: "— nessuno —" }, ...data.creativi
    .filter((c) => c.tipo === "adset"
      && (!ed || !ed.draft.campagnaId || c.campagnaId === ed.draft.campagnaId)
      && (!ed || !ed.draft.clienteId || c.clienteId === ed.draft.clienteId)
      && (!ed || c.id !== ed.draft.id))
    .map((c) => ({ v: c.id, l: c.nome }))];

  /* I modelli predefiniti più quelli che ti sei costruito */
  const optCategorie = [{ v: "", l: "— senza categoria —" }, ...(data.categorie || []).map((c) => ({ v: c.id, l: c.nome }))];
  const optCliente = [{ v: "", l: "— tutta l'agenzia —" }, ...attivi.map((z) => ({ v: z.id, l: z.nome }))];

  const optModelli = [
    ...MODELLI_FUNNEL,
    ...(data.modelli || []).map((x) => ({ v: `p:${x.id}`, l: `${x.nome} · tuo modello`, passi: x.passi })),
  ];

  const optFunnel = [{ v: "", l: "— nessuno —" }, ...(attivo ? S.funnel : data.funnel).map((f) => ({
    v: f.id, l: attivo ? f.nome : `${nomeCli(f.clienteId)} · ${f.nome}`,
  }))];
  const optCampScelta = [{ v: "", l: "— nessuna —" }, ...(attivo ? S.campagne : m.campagne).map((k) => ({
    v: k.id, l: attivo ? k.nome : `${nomeCli(k.clienteId)} · ${k.nome}`,
  }))];
  const optContCli = [{ v: "", l: "— nessuno —" }, ...S.contatti.map((c) => ({ v: c.id, l: `${c.nome}${c.azienda ? " · " + c.azienda : ""}` }))];

  /* ---------- grafici ---------- */
  const sommaVendite = (l) => CAMPI_VENDITA.reduce((o, k) => ({ ...o, [k]: l.reduce((s, r) => s + (r[k] || 0), 0) }), {});
  const chiusureVista = attivo ? S.chiusure : data.chiusure;

  /* Se per quel mese esistono schede chiusura, sono loro a fare testo:
     chiusure, contratto e incassato arrivano da lì invece che a mano. */
  /* Se per quel mese ci sono giornate caricate dai venditori, i numeri di attività
     arrivano da lì: una riga al giorno invece di un totale scritto a mano. */
  const conGiornate = (r) => {
    const gs = data.giornate.filter((g) => g.clienteId === r.clienteId && mese(g.data || "") === r.periodo);
    if (!gs.length) return r;
    const somma = (k) => gs.reduce((x, g) => x + (g[k] || 0), 0);
    const conOre = gs.filter((g) => g.oreContatto > 0);
    return {
      ...r, daGiornate: gs.length,
      lead: somma("lead"), chiamate: somma("chiamate"),
      discoveryFissate: somma("discoveryFissate"), discoverySvolte: somma("discoverySvolte"),
      demoFissate: somma("demoFissate"), demoSvolte: somma("demoSvolte"), offerte: somma("offerte"),
      oreContatto: conOre.length ? conOre.reduce((x, g) => x + g.oreContatto, 0) / conOre.length : (r.oreContatto || 0),
    };
  };

  const conChiusure = (r) => {
    const cs = data.chiusure.filter((c) => c.clienteId === r.clienteId && mese(c.data || "") === r.periodo);
    if (!cs.length) return { ...r, auto: false };
    return {
      ...r, auto: true, chiusure: cs.length,
      valoreContratto: cs.reduce((s, c) => s + (c.valoreContratto || 0), 0),
      incassato: cs.reduce((s, c) => s + (c.incassato || 0), 0),
    };
  };
  const venditeVista = (attivo ? S.vendite : data.vendite).map(conGiornate).map(conChiusure);

  const METRICHE_C = [
    { v: "spesa", l: "Spesa", f: (g) => g.sp, fm: eur },
    { v: "ricavi", l: "Ricavi", f: (g) => g.rc, fm: eur },
    { v: "conversioni", l: "Conversioni", f: (g) => g.cv, fm: int },
    { v: "click", l: "Click", f: (g) => g.cl, fm: int },
    { v: "impression", l: "Impression", f: (g) => g.imp, fm: int },
    { v: "roas", l: "ROAS", f: (g) => (g.sp ? g.rc / g.sp : 0), fm: (x) => `${fmt(x)}×`, rap: true },
    { v: "cpa", l: "CPA", f: (g) => (g.cv ? g.sp / g.cv : 0), fm: eur, rap: true },
    { v: "ctr", l: "CTR", f: (g) => (g.imp ? (g.cl / g.imp) * 100 : 0), fm: (x) => pctv(x), rap: true },
    { v: "cpc", l: "CPC", f: (g) => (g.cl ? g.sp / g.cl : 0), fm: eur, rap: true },
    { v: "cpm", l: "CPM", f: (g) => (g.imp ? (g.sp / g.imp) * 1000 : 0), fm: eur, rap: true },
  ];
  const METRICHE_V = [
    { v: "chiusure", l: "Chiusure", f: (g) => g.chiusure, fm: int },
    { v: "valoreContratto", l: "Valore contratto", f: (g) => g.valoreContratto, fm: eur },
    { v: "incassato", l: "Importo incassato", f: (g) => g.incassato, fm: eur },
    { v: "lead", l: "Lead", f: (g) => g.lead, fm: int },
    { v: "chiamate", l: "Chiamate", f: (g) => g.chiamate, fm: int },
    { v: "discoverySvolte", l: "Discovery svolte", f: (g) => g.discoverySvolte, fm: int },
    { v: "demoSvolte", l: "Demo svolte", f: (g) => g.demoSvolte, fm: int },
    { v: "showDisco", l: "Show up discovery", f: (g) => (g.discoveryFissate ? (g.discoverySvolte / g.discoveryFissate) * 100 : 0), fm: (x) => pctv(x), rap: true },
    { v: "showDemo", l: "Show up demo", f: (g) => (g.demoFissate ? (g.demoSvolte / g.demoFissate) * 100 : 0), fm: (x) => pctv(x), rap: true },
    { v: "chiusura", l: "Tasso di chiusura", f: (g) => (g.demoSvolte ? (g.chiusure / g.demoSvolte) * 100 : 0), fm: (x) => pctv(x), rap: true },
    { v: "leadChiusura", l: "Lead → chiusura", f: (g) => (g.lead ? (g.chiusure / g.lead) * 100 : 0), fm: (x) => pctv(x), rap: true },
    { v: "incasso", l: "Quota incassata", f: (g) => (g.valoreContratto ? (g.incassato / g.valoreContratto) * 100 : 0), fm: (x) => pctv(x), rap: true },
    { v: "ticket", l: "Contratto medio", f: (g) => (g.chiusure ? g.valoreContratto / g.chiusure : 0), fm: eur, rap: true },
    { v: "cac", l: "Costo per chiusura", f: (g) => (g.chiusure ? g.spesa / g.chiusure : 0), fm: eur, rap: true },
  ];
  const METRICHE_CH = [
    { v: "numero", l: "Numero di chiusure", f: (l) => l.length, fm: int },
    { v: "contratto", l: "Valore contratto", f: (l) => l.reduce((s, c) => s + (c.valoreContratto || 0), 0), fm: eur },
    { v: "incassato", l: "Importo incassato", f: (l) => l.reduce((s, c) => s + (c.incassato || 0), 0), fm: eur },
    { v: "residuo", l: "Da incassare", f: (l) => l.reduce((s, c) => s + ((c.valoreContratto || 0) - (c.incassato || 0)), 0), fm: eur },
    { v: "medio", l: "Contratto medio", f: (l) => (l.length ? l.reduce((s, c) => s + (c.valoreContratto || 0), 0) / l.length : 0), fm: eur, rap: true },
  ];
  const GRUPPI_CH = [
    { v: "funnel", l: "Per funnel" }, { v: "campagna", l: "Per campagna" }, { v: "adset", l: "Per ad set" },
    { v: "annuncio", l: "Per annuncio" }, { v: "mese", l: "Per mese" },
    ...(attivo ? [] : [{ v: "cliente", l: "Per cliente" }]),
  ];
  const GRUPPI_C = [
    { v: "campagna", l: "Per campagna" }, { v: "piattaforma", l: "Per piattaforma" },
    { v: "obiettivo", l: "Per obiettivo" }, ...(attivo ? [] : [{ v: "cliente", l: "Per cliente" }]),
  ];
  const GRUPPI_V = [
    { v: "fasi", l: "Fasi dell'imbuto" }, { v: "periodo", l: "Per mese" },
    ...(attivo ? [] : [{ v: "cliente", l: "Per cliente" }]),
  ];
  const METRICHE_FASI = [{ v: "numeri", l: "Numeri assoluti" }, { v: "tassi", l: "Tassi di passaggio" }];

  const metricheDisponibili = graf.ambito === "campagne" ? METRICHE_C
    : graf.ambito === "chiusure" ? METRICHE_CH
    : graf.gruppo === "fasi" ? METRICHE_FASI : METRICHE_V;
  const gruppiDisponibili = graf.ambito === "campagne" ? GRUPPI_C : graf.ambito === "chiusure" ? GRUPPI_CH : GRUPPI_V;

  const cambiaGraf = (patch) => {
    let g = { ...graf, ...patch };
    if (patch.ambito) {
      g.gruppo = { campagne: "campagna", vendite: "fasi", chiusure: "funnel" }[patch.ambito];
      g.metrica = { campagne: "spesa", vendite: "numeri", chiusure: "contratto" }[patch.ambito];
    }
    if (patch.gruppo && graf.ambito === "vendite") {
      g.metrica = patch.gruppo === "fasi" ? "numeri" : "chiusure";
    }
    setGraf(g);
  };

  const costruisci = () => {
    if (graf.ambito === "campagne") {
      const base = attivo ? S.campagne : m.campagne;
      const chiave = {
        campagna: (k) => k.nome,
        piattaforma: (k) => et(PIATTAFORME, k.piattaforma),
        obiettivo: (k) => et(OBIETTIVI, k.obiettivo),
        cliente: (k) => nomeCli(k.clienteId),
      }[graf.gruppo] || ((k) => k.nome);
      const mappa = new Map();
      base.forEach((k) => {
        const key = chiave(k);
        const g = mappa.get(key) || { sp: 0, rc: 0, cv: 0, cl: 0, imp: 0 };
        g.sp += k.d.sp; g.rc += k.d.rc; g.cv += k.d.cv; g.cl += k.d.cl; g.imp += k.d.imp;
        mappa.set(key, g);
      });
      const me = METRICHE_C.find((x) => x.v === graf.metrica) || METRICHE_C[0];
      const dati = [...mappa.entries()].map(([k, g]) => ({ k, v: me.f(g) })).sort((a, b) => b.v - a.v).slice(0, 12);
      return { dati, fmtv: me.fm, rap: !!me.rap, titolo: `${me.l} · ${et(gruppiDisponibili, graf.gruppo).toLowerCase()}` };
    }

    if (graf.ambito === "chiusure") {
      const nomeCampagna = (id) => (data.campagne.find((k) => k.id === id) || {}).nome || "senza campagna";
      const chiave = {
        funnel: (c) => et(FUNNEL_PROV, c.funnel),
        campagna: (c) => nomeCampagna(c.campagnaId),
        adset: (c) => c.adset || "ad set non indicato",
        annuncio: (c) => c.annuncio || "annuncio non indicato",
        mese: (c) => mese(c.data || ""),
        cliente: (c) => nomeCli(c.clienteId),
      }[graf.gruppo] || ((c) => et(FUNNEL_PROV, c.funnel));
      const mappa = new Map();
      chiusureVista.forEach((c) => mappa.set(chiave(c), [...(mappa.get(chiave(c)) || []), c]));
      const me = METRICHE_CH.find((x) => x.v === graf.metrica) || METRICHE_CH[0];
      let dati = [...mappa.entries()].map(([k, l]) => ({ k: graf.gruppo === "mese" ? meseIt(k) : k, ord: k, v: me.f(l) }));
      dati = graf.gruppo === "mese" ? dati.sort((a, b) => (a.ord < b.ord ? -1 : 1)) : dati.sort((a, b) => b.v - a.v).slice(0, 12);
      return { dati, fmtv: me.fm, rap: !!me.rap, titolo: `${me.l} · ${et(gruppiDisponibili, graf.gruppo).toLowerCase()}` };
    }

    if (graf.gruppo === "fasi") {
      const t = sommaVendite(venditeVista);
      if (graf.metrica === "tassi") {
        const dati = FASI_VENDITA.slice(1).map((f, i) => {
          const prec = FASI_VENDITA[i].k;
          return { k: `${FASI_VENDITA[i].l} → ${f.l}`, v: t[prec] ? (t[f.k] / t[prec]) * 100 : 0 };
        });
        return { dati, fmtv: (x) => pctv(x), rap: true, titolo: "Tassi di passaggio tra le fasi" };
      }
      const dati = FASI_VENDITA.map((f) => ({ k: f.l, v: t[f.k] || 0 }));
      return { dati, fmtv: int, rap: false, titolo: "Percorso commerciale, dalle prime chiamate alle chiusure" };
    }

    const mappa = new Map();
    venditeVista.forEach((r) => {
      const key = graf.gruppo === "cliente" ? nomeCli(r.clienteId) : r.periodo;
      mappa.set(key, [...(mappa.get(key) || []), r]);
    });
    const me = METRICHE_V.find((x) => x.v === graf.metrica) || METRICHE_V[0];
    let dati = [...mappa.entries()].map(([k, l]) => ({ k: graf.gruppo === "periodo" ? meseIt(k) : k, ord: k, v: me.f(sommaVendite(l)) }));
    dati = graf.gruppo === "periodo" ? dati.sort((a, b) => (a.ord < b.ord ? -1 : 1)) : dati.sort((a, b) => b.v - a.v);
    return { dati, fmtv: me.fm, rap: !!me.rap, titolo: `${me.l} · ${graf.gruppo === "periodo" ? "mese per mese" : "per cliente"}` };
  };

  /* I totali di una campagna con finestre compilate nascono dalla somma delle
     finestre: quello che c'è scritto nella scheda viene ignorato. Finché era
     invisibile si poteva correggere un numero e non vedere cambiare niente,
     quindi qui calcolo la somma e la mostro al posto del campo libero. */
  const fasiDellaScheda = (() => {
    if (!ed || ed.tipo !== "campagne" || !ed.draft || !ed.draft.id) return null;
    const CH = ["spesa", "impression", "click", "clickLink", "lead", "conversioni", "ricavi"];
    const g = { n: 0, fq: 0 };
    CH.forEach((c) => { g[c] = 0; });
    (data.periodi || []).filter((pz) => pz.campagnaId === ed.draft.id).forEach((pz) => {
      g.n += 1;
      CH.forEach((c) => { g[c] += pz[c] || 0; });
      g.fq += (pz.frequenza || 0) * (pz.impression || 0);
    });
    if (!g.n) return null;
    if (g.spesa + g.impression + g.click + g.conversioni + g.ricavi === 0) return null;
    g.frequenza = g.impression ? g.fq / g.impression : 0;
    return g;
  })();
  /* Un campo che non si può cambiare mostra il valore vero e dice dove si cambia */
  const daFinestre = (k, label, extra = {}) => (fasiDellaScheda
    ? { k, label, type: "num", ...extra, bloccato: perInput(Math.round((fasiDellaScheda[k] || 0) * 100) / 100),
      hint: `somma delle ${fasiDellaScheda.n} finestre` }
    : { k, label, type: "num", ...extra });

  const CAMPI = {
    agenzia: [
      { k: "nome", label: "Nome agenzia", w2: true },
      { k: "sottotitolo", label: "Sottotitolo" },
      { k: "costoOrario", label: "Costo orario del tuo tempo", type: "num", w2: true, hint: "serve a calcolare il margine per cliente" },
      { k: "logo", label: "Logo", type: "logo", wide: true },
      { k: "lingua", label: "Lingua dell'interfaccia", type: "select", w2: true,
        opts: LINGUE.map((x) => ({ v: x.v, l: x.l })) },
      { k: "valuta", label: "Valuta predefinita", type: "select", w2: true,
        opts: VALUTE.map((c) => ({ v: c, l: c })), hint: "vale per i clienti che non ne hanno una propria" },
    ],
    aree: [{ k: "nome", label: "Area di lavoro", wide: true }],
    clienti: [
      { k: "nome", label: "Cliente", w2: true },
      { k: "stato", label: "Stato", type: "select", opts: STATI_CLIENTE },
      { k: "cognome", label: "Cognome per l'ordinamento", w2: true,
        hint: "solo se il nome non basta a dedurlo, per esempio quando è scritto al contrario" },
      { k: "padreId", label: "Sottocliente di", type: "select", wide: true,
        opts: [{ v: "", l: "— cliente principale —" }, ...data.clienti
          .filter((z) => !ed || (z.id !== ed.draft.id && !discendenti(ed.draft.id || "").includes(z.id)))
          .map((z) => ({ v: z.id, l: z.padreId ? `${nomeCli(z.padreId)} › ${z.nome}` : z.nome }))] },
      { k: "logo", label: "Logo del cliente", type: "logo", wide: true },
      { k: "settore", label: "Settore", w2: true },
      { k: "citta", label: "Sede" },
      { k: "creato", label: "In gestione dal", type: "date", hint: "la data da cui lo segui, non quella di inserimento" },
      { k: "dataContratto", label: "Data contratto", type: "date", hint: "quando è stato firmato" },
      { k: "accountAdv", label: "Account pubblicitario", w2: true, hint: "l'ID di Meta, quello che inizia per act_" },
      { k: "valuta", label: "Valuta", type: "select", w2: true,
        opts: [{ v: "", l: "— quella dell'agenzia (o del cliente madre) —" }, ...VALUTE.map((c) => ({ v: c, l: c }))],
        hint: "se questo cliente fattura in un'altra valuta" },
      { k: "fee", label: "Compenso mensile", type: "num" },
      { k: "valutaFee", label: "Valuta del compenso", type: "select",
        opts: [{ v: "", l: "— come la valuta del cliente —" }, ...VALUTE.map((c) => ({ v: c, l: c }))],
        hint: "se ti paga in una valuta diversa da quella delle campagne" },
      { k: "budgetMese", label: "Budget adv mensile", type: "num" },
      { grp: "Economia del cliente", hint: "servono al pareggio e alla soglia di valutazione" },
      { k: "margineLordo", label: "Margine lordo (%)", type: "num", ph: "42" },
      { k: "resi", label: "Resi e insoluti (%)", type: "num", ph: "6" },
      { k: "aov", label: "Valore medio ordine", type: "num", ph: "180", hint: "la soglia di giudizio è il doppio" },
      { grp: "Dati di fatturazione" },
      { k: "ragioneSociale", label: "Ragione sociale", w2: true },
      { k: "piva", label: "Partita IVA" },
      { k: "codiceFiscale", label: "Codice fiscale" },
      { k: "indirizzo", label: "Indirizzo (sede legale)", w2: true },
      { k: "cap", label: "CAP" },
      { k: "cittaFatt", label: "Città" },
      { k: "provincia", label: "Provincia" },
      { k: "paese", label: "Paese" },
      { k: "sdi", label: "Codice destinatario (SDI)" },
      { k: "pec", label: "PEC", w2: true },
      { k: "emailAmm", label: "Email amministrazione", w2: true },
      { k: "referenteAmm", label: "Referente amministrativo" },
      { k: "termini", label: "Termini di pagamento", type: "select", opts: TERMINI_PAGAMENTO },
      { k: "metodo", label: "Metodo", type: "select", opts: METODI_PAGAMENTO },
      { k: "note", label: "Note e vincoli", type: "textarea", wide: true },
    ],
    contatti: [
      { k: "tipo", label: "Tipo", type: "select", opts: TIPI_CONTATTO },
      { k: "nome", label: "Nome e cognome", w2: true },
      { k: "azienda", label: "Azienda", w2: true },
      { k: "ruolo", label: "Ruolo" },
      { k: "email", label: "Email", w2: true },
      { k: "telefono", label: "Telefono" },
      { k: "stato", label: "Stato", type: "select", opts: STATI_LEAD },
      { k: "campagnaId", label: "Campagna di provenienza", type: "select", opts: optCampCli, w2: true },
      { k: "note", label: "Note", type: "textarea", wide: true },
    ],
    trattative: [
      { k: "titolo", label: "Trattativa", wide: true },
      { k: "contattoId", label: "Contatto", type: "select", opts: optContCli, w2: true },
      { k: "valore", label: "Valore", type: "num" },
      { k: "fase", label: "Fase", type: "select", opts: FASI },
      { k: "chiusura", label: "Chiusura prevista", type: "date" },
      { k: "campagnaId", label: "Campagna di origine", type: "select", opts: optCampCli },
      { k: "note", label: "Note", type: "textarea", wide: true },
    ],
    campagne: [
      /* A livello agenzia nessun cliente è aperto: senza questo campo la campagna nascerebbe orfana */
      ...(cliAtt ? [] : [{ k: "clienteId", label: "Cliente", type: "select", wide: true,
        opts: [{ v: "", l: "— scegli il cliente —" }, ...data.clienti.filter((z) => z.archiviato !== "si").map((z) => ({ v: z.id, l: z.nome }))] }]),
      { k: "nome", label: "Campagna", wide: true },
      { k: "piattaforma", label: "Piattaforma", type: "select", opts: PIATTAFORME },
      { k: "obiettivo", label: "Obiettivo", type: "select", opts: OBIETTIVI },
      { k: "stato", label: "Stato", type: "select", opts: STATI_CAMPAGNA },
      { k: "linea", label: "Linea di business", w2: true, lista: "linee",
        hint: "campagne di linee diverse non vengono sommate" },
      { k: "funnelId", label: "Funnel di destinazione", type: "select", opts: optFunnel, w2: true },
      { k: "archiviata", label: "Archiviata", type: "select", opts: SN, hint: "fuori da medie e totali" },
      { k: "inizio", label: "Data di partenza", type: "date" },
      { k: "fine", label: "Data di fine", type: "date", hint: "da compilare quando la spegni o la concludi" },
      { k: "ultimaModifica", label: "Ultima modifica", type: "date", hint: "da qui partono i tre giorni di attesa prima di rigiudicare" },
      { k: "notaModifica", label: "Cosa hai cambiato", w2: true, hint: "per esempio: spento AD1, alzato il budget" },
      { k: "budgetGiorno", label: "Budget giornaliero", type: "num", hint: "quello attuale" },
      { grp: "Variazioni di budget", hint: "ogni cambio con la sua data, senza limite" },
      { k: "__budget", label: "Storico del giornaliero", type: "budget", wide: true },
      { grp: "Fasi", hint: "Test 3 giorni · Matura 7 giorni · poi una settimana alla volta" },
      { k: "__fasi", label: "Finestre della campagna", type: "fasi", wide: true },
      { k: "budget", label: "Budget totale del periodo", type: "num", hint: "facoltativo" },
      { grp: "Dati del periodo",
        hint: fasiDellaScheda
          ? `Questi totali li calcola l'app sommando le ${fasiDellaScheda.n} finestre compilate, quindi qui non si modificano: apri il pulsante Fasi nella tabella e scrivi i numeri nella finestra giusta. Valore contratto e incassato restano tuoi.`
          : "puoi usare la virgola: 1.234,56" },
      daFinestre("spesa", "Speso"),
      daFinestre("impression", "Impression"),
      daFinestre("click", "Clic totali"),
      daFinestre("clickLink", "Clic sul link", { hint: "quelli che contano" }),
      daFinestre("lead", "Lead", { hint: "per il costo per lead" }),
      /* Conversioni e ricavi restano scrivibili anche con le finestre: sono l'esito
         commerciale — appuntamenti, vendite — non un numero che arriva da Meta.
         In una campagna a contatti il lead non è una conversione. */
      { k: "conversioni", label: "Conversioni", type: "num",
        hint: fasiDellaScheda
          ? `esito a valle del lead: appuntamenti, vendite. Le finestre ne sommano ${int(fasiDellaScheda.conversioni || 0)}; se scrivi un numero diverso vale il tuo`
          : "esito a valle del lead: appuntamenti, vendite" },
      { k: "ricavi", label: "Ricavi tracciati", type: "num",
        hint: fasiDellaScheda
          ? `le finestre sommano ${conValuta(valutaCliente(data, ed.draft.clienteId), () => eur(fasiDellaScheda.ricavi || 0))}; se scrivi un numero diverso vale il tuo`
          : "quelli dichiarati dalla piattaforma" },
      { k: "valoreContratto", label: "Valore contratto", type: "num", hint: "firmato" },
      { k: "incassato", label: "Incassato", type: "num", hint: "entrato davvero" },
      daFinestre("frequenza", "Frequenza", { ph: "0,00" }),
      { grp: "Benchmark", hint: "0 = metrica non controllata" },
      { k: "roasMin", label: "ROAS minimo (×)", type: "num", ph: "0,00" },
      { k: "cplMax", label: "CPL massimo", type: "num", ph: "0,00", hint: "nelle campagne a contatti è il KPI del semaforo" },
      { k: "cpaMax", label: "CPA massimo", type: "num", ph: "0,00" },
      { k: "ctrMin", label: "CTR minimo (%)", type: "num", ph: "0,81" },
      { k: "freqMax", label: "Frequenza massima", type: "num", ph: "0,00" },
      { k: "cpmMax", label: "CPM massimo", type: "num", ph: "0,00" },
    ],
    vendite: [
      { k: "clienteId", label: "Cliente", type: "select", opts: optCliTutti, w2: true },
      { k: "periodo", label: "Mese", type: "month" },
      { grp: "Percorso commerciale", hint: "i numeri di questo periodo" },
      { k: "lead", label: "Lead ricevuti", type: "num" },
      { k: "chiamate", label: "Chiamate fatte", type: "num" },
      { k: "discoveryFissate", label: "Discovery fissate", type: "num" },
      { k: "discoverySvolte", label: "Discovery svolte", type: "num" },
      { k: "demoFissate", label: "Demo fissate", type: "num" },
      { k: "demoSvolte", label: "Demo svolte", type: "num" },
      { k: "offerte", label: "Offerte presentate", type: "num" },
      { k: "chiusure", label: "Chiusure", type: "num", hint: "ignorato se ci sono schede chiusura" },
      { k: "valoreContratto", label: "Valore contratto", type: "num" },
      { k: "incassato", label: "Importo incassato", type: "num" },
      { k: "spesa", label: "Spesa adv del periodo", type: "num" },
      { k: "convPiattaforma", label: "Conversioni dichiarate dalle piattaforme", type: "num", hint: "per misurare lo scarto con i lead reali" },
      { grp: "Capacità commerciale", hint: "dove si strozza l'imbuto" },
      { k: "slot", label: "Slot discovery disponibili", type: "num", hint: "quante ne potevi gestire" },
      { k: "oreContatto", label: "Ore al primo contatto", type: "num", ph: "4", hint: "tempo medio lead → prima chiamata" },
    ],
    chiusure: [
      { k: "nome", label: "Cliente chiuso", w2: true },
      { k: "data", label: "Data chiusura", type: "date" },
      { k: "clienteId", label: "Account", type: "select", opts: optCliTutti, w2: true },
      { k: "closer", label: "Closer" },
      { k: "azienda", label: "Azienda", w2: true },
      { k: "telefono", label: "Telefono" },
      { k: "email", label: "Email", wide: true },
      { grp: "Provenienza", hint: "da dove è arrivato davvero" },
      { k: "funnel", label: "Funnel", type: "select", opts: FUNNEL_PROV },
      { k: "campagnaId", label: "Campagna", type: "select", opts: optCampScelta, w2: true },
      { k: "adset", label: "Ad set", w2: true },
      { k: "annuncio", label: "Annuncio", w2: true },
      { grp: "Importi" },
      { k: "valoreContratto", label: "Valore contratto", type: "num" },
      { k: "incassato", label: "Importo incassato", type: "num" },
      { k: "note", label: "Note su pagamenti e accordi", type: "textarea", wide: true },
    ],
    annotazioni: [
      { k: "data", label: "Data" , type: "date" },
      { k: "tipo", label: "Tipo di intervento", type: "select", opts: TIPI_NOTA },
      { k: "clienteId", label: "Cliente", type: "select", opts: optCliTutti },
      { k: "campagnaId", label: "Campagna", type: "select", opts: optCampScelta, wide: true },
      { k: "testo", label: "Cosa hai cambiato e perché", type: "textarea", wide: true },
      { k: "avviaAttesa", label: "È un intervento: avvia i tre giorni di attesa", type: "flag", wide: true,
        hint: "il semaforo sospende il giudizio su quella campagna fino al controllo" },
    ],
    esperimenti: [
      { k: "ipotesi", label: "Ipotesi", wide: true, hint: "cosa credi che succeda, in una frase" },
      { k: "clienteId", label: "Cliente", type: "select", opts: optCliTutti },
      { k: "variabile", label: "Cosa cambi", w2: true },
      { k: "metrica", label: "Metrica di giudizio", w2: true },
      { k: "stato", label: "Esito", type: "select", opts: STATI_ESPERIMENTO },
      { k: "inizio", label: "Inizio", type: "date" },
      { k: "fine", label: "Fine", type: "date" },
      { k: "risultato", label: "Risultato misurato", wide: true },
      { k: "conclusione", label: "Cosa ne ricavi per il futuro", type: "textarea", wide: true },
    ],
    eventi: [
      { k: "titolo", label: "Titolo", wide: true },
      { k: "data", label: "Dal", type: "date" },
      { k: "dataFine", label: "Al", type: "date", hint: "lascia vuoto se dura un giorno" },
      { k: "ora", label: "Ora", type: "time", w2: true, hint: "vuoto significa tutto il giorno" },
      { k: "durata", label: "Durata in ore", num: true, w2: true, hint: "serve solo alla vista giornaliera" },
      { k: "categoriaId", label: "Categoria", type: "select", opts: optCategorie, w2: true },
      { k: "clienteId", label: "Cliente", type: "select", opts: optCliente },
      { k: "note", label: "Note", type: "textarea", wide: true },
    ],
    categorie: [
      { k: "nome", label: "Nome della categoria", w2: true },
      { k: "colore", label: "Colore", type: "colore", hint: "sarà il colore sul calendario" },
    ],
    modelli: [
      { k: "nome", label: "Nome del modello", wide: true, hint: "come lo ritroverai nel modulo del funnel" },
      { grp: "Passi", hint: "la sequenza, senza numeri: quelli li metterai nel funnel" },
      { k: "passi", label: "Sequenza", type: "passi", wide: true },
    ],
    funnel: [
      { k: "nome", label: "Nome del funnel", w2: true },
      { k: "obiettivo", label: "Obiettivo", type: "select", opts: OBIETTIVI_FUNNEL },
      { k: "clienteId", label: "Cliente", type: "select", opts: optCliTutti, w2: true },
      { k: "stato", label: "Stato", type: "select", opts: [{ v: "attiva", l: "Attivo" }, { v: "test", l: "In test" }, { v: "pausa", l: "In pausa" }] },
      { k: "modello", label: "Parti da un modello", type: "select", opts: optModelli, wide: true, hint: "sceglierlo riscrive i passi" },
      { k: "__campagne", label: "Campagne che portano traffico", type: "campagne", wide: true,
        hint: "spunta quelle che alimentano questo percorso: puoi sceglierne una, alcune o tutte" },
      { k: "spesa", label: "Spesa", type: "num" },
      { k: "valoreMedio", label: "Valore medio a conversione", type: "num", w2: true },
      { grp: "Passi", hint: "visite = chi arriva · passano = chi va allo step dopo" },
      { k: "passi", label: "Sequenza", type: "passi", wide: true },
      { grp: "Pagina di destinazione", hint: "genera un file HTML da caricare dove vuoi" },
      { k: "pTitolo", label: "Titolo", wide: true },
      { k: "pSottotitolo", label: "Sottotitolo", type: "textarea", wide: true },
      { k: "pPunti", label: "Punti elenco (uno per riga)", type: "textarea", wide: true },
      { k: "pCta", label: "Testo del pulsante", w2: true },
      { k: "pWebhook", label: "Webhook che riceve i moduli", wide: true, hint: "lo stesso tipo di URL delle automazioni" },
      { k: "note", label: "Note", type: "textarea", wide: true },
    ],
    file: [
      { k: "titolo", label: "Titolo", w2: true },
      { k: "categoria", label: "Categoria", type: "select", opts: CATEGORIE_FILE },
      { k: "clienteId", label: "Cliente", type: "select", opts: [{ v: "", l: "— materiale dell'agenzia —" }, ...data.clienti.map((z) => ({ v: z.id, l: z.nome }))], w2: true },
      { k: "tipo", label: "Tipo", type: "select", opts: TIPI_FILE },
      { k: "url", label: "Collegamento (Drive, Dropbox, Notion…)", wide: true },
      { k: "tag", label: "Etichette", w2: true },
      { k: "aggiornato", label: "Aggiornato il", type: "date" },
      { k: "testo", label: "Contenuto del documento", type: "textarea", wide: true },
    ],
    messaggio: [
      { grp: "Il mercato", hint: "chi è, e cosa gli toglie il sonno" },
      { k: "mercato", label: "Chi è il mercato", w2: true, hint: "es. titolari di aziende manifatturiere nelle Marche" },
      { k: "consapevolezza", label: "Consapevolezza", type: "select", opts: CONSAPEVOLEZZA },
      { k: "avatarChi", label: "Com'è fatto tipicamente", type: "textarea", wide: true },
      { k: "avatarCredenze", label: "Cosa conta davvero per lui, quali convinzioni ha", type: "textarea", wide: true },
      { k: "avatarMedia", label: "Dove lo trovi e cosa consuma", type: "textarea", wide: true },
      { grp: "Dolore e desiderio" },
      { k: "dolore", label: "Il dolore che risolvi o il guadagno che dai", type: "textarea", wide: true },
      { k: "doloriCompratore", label: "Dolori di chi compra davvero", type: "textarea", w2: true, hint: "attuali e urgenti" },
      { k: "doloriCuriosi", label: "Dolori dei curiosi che non comprano", type: "textarea", hint: "veri ma non urgenti" },
      { k: "desiderio", label: "Il desiderio di massa", w2: true },
      { k: "tensione", label: "La tensione emotiva da togliere", hint: "paura, vergogna, fatica" },
      { k: "futuro", label: "Come sarà la sua vita dopo il risultato", type: "textarea", wide: true },
      { grp: "Offerta e differenza" },
      { k: "cosaVendo", label: "Cosa vendi", w2: true },
      { k: "offerta", label: "L'offerta in una riga" },
      { k: "soluzione", label: "La tua soluzione unica al dolore", type: "textarea", wide: true },
      { k: "benefici", label: "I tre grandi benefici (uno per riga)", type: "textarea", wide: true },
      { k: "obiezioni", label: "Le tre grandi obiezioni (una per riga)", type: "textarea", wide: true },
      { k: "missione", label: "La missione dichiarata al mercato", type: "textarea", wide: true },
      { k: "usp", label: "Cosa puoi dire solo tu", type: "textarea", wide: true, hint: "se lo può dire anche un concorrente, non è una USP" },
    ],
    periodi: [
      { k: "tipo", label: "Fase", type: "select", opts: TIPI_PERIODO },
      { k: "inizio", label: "Dal", type: "date" },
      { k: "fine", label: "Al", type: "date" },
      { grp: "Dati del periodo", hint: "solo quelli di questa finestra, non cumulati" },
      { k: "spesa", label: "Speso", type: "num" },
      { k: "impression", label: "Impression", type: "num" },
      { k: "click", label: "Clic totali", type: "num" },
      { k: "clickLink", label: "Clic sul link", type: "num", hint: "quelli che contano" },
      { k: "lead", label: "Lead", type: "num", hint: "per il costo per lead" },
      { k: "conversioni", label: "Conversioni", type: "num" },
      { k: "ricavi", label: "Ricavi tracciati", type: "num", hint: "quelli dichiarati dalla piattaforma" },
      { k: "valoreContratto", label: "Valore contratto", type: "num", hint: "firmato" },
      { k: "incassato", label: "Incassato", type: "num", hint: "entrato davvero" },
      { k: "frequenza", label: "Frequenza", type: "num", ph: "0,00" },
      { k: "note", label: "Cosa è successo in questa finestra", type: "textarea", wide: true },
    ],
    creativi: [
      { k: "nome", label: "Nome", w2: true },
      { k: "tipo", label: "Livello", type: "select", opts: TIPI_CREATIVO },
      { k: "campagnaId", label: "Campagna", type: "select", opts: optCampScelta, w2: true },
      { k: "stato", label: "Stato", type: "select", opts: STATI_CAMPAGNA },
      { k: "adsetId", label: "Ad set di appartenenza", type: "select", opts: optAdset, wide: true,
        hint: "serve per gli annunci: li raggruppa sotto il loro ad set" },
      { k: "stadio", label: "Fase del messaggio", type: "select", opts: [{ v: "", l: "— non indicata —" }, ...STADI], w2: true, hint: "a quale momento del percorso d\'acquisto parla" },
      { grp: "Dati del periodo", hint: "gli stessi che copi dalla piattaforma" },
      { k: "spesa", label: "Speso", type: "num" },
      { k: "impression", label: "Impression", type: "num" },
      { k: "click", label: "Clic totali", type: "num" },
      { k: "clickLink", label: "Clic sul link", type: "num", hint: "quelli che contano" },
      { k: "lead", label: "Lead", type: "num", hint: "per il costo per lead" },
      { k: "conversioni", label: "Conversioni", type: "num" },
      { k: "ricavi", label: "Ricavi tracciati", type: "num", hint: "quelli dichiarati dalla piattaforma" },
      { k: "valoreContratto", label: "Valore contratto", type: "num", hint: "firmato" },
      { k: "incassato", label: "Incassato", type: "num", hint: "entrato davvero" },
      { k: "frequenza", label: "Frequenza", type: "num", ph: "0,00" },
      { k: "note", label: "Note", type: "textarea", wide: true },
    ],
    automazioni: [
      { k: "nome", label: "Nome della regola", w2: true },
      { k: "attiva", label: "Stato", type: "select", opts: [{ v: "si", l: "Attiva" }, { v: "no", l: "In pausa" }] },
      { k: "clienteId", label: "Si applica a", type: "select", opts: [{ v: "", l: "Tutti i clienti" }, ...data.clienti.map((z) => ({ v: z.id, l: z.nome }))], wide: true },
      { grp: "Quando", hint: "cosa fa scattare la regola" },
      { k: "trigger", label: "Innesco", type: "select", opts: TRIGGER, wide: true },
      { k: "filtroAvviso", label: "Solo per la metrica", type: "select", opts: FILTRI_AVVISO, w2: true, hint: "vale per l'innesco su avviso" },
      { k: "soloCritici", label: "Solo avvisi critici", type: "select", opts: SN },
      { k: "metrica", label: "Metrica", type: "select", opts: METRICHE_TRIGGER, hint: "per l'innesco su soglia" },
      { k: "operatore", label: "Condizione", type: "select", opts: OPERATORI },
      { k: "valore", label: "Valore soglia", type: "num", hint: "per budget è una percentuale" },
      { k: "giorno", label: "Giorno del mese", type: "num", hint: "per il controllo dati mancanti" },
      { grp: "Cosa fare" },
      { k: "azTask", label: "Crea un task", type: "select", opts: SN },
      { k: "taskTitolo", label: "Titolo del task", w2: true },
      { k: "taskArea", label: "Area", type: "select", opts: optAree },
      { k: "taskPriorita", label: "Priorità", type: "select", opts: PRIORITA },
      { k: "taskGiorni", label: "Scadenza tra (giorni)", type: "num" },
      { k: "azWebhook", label: "Chiama un webhook", type: "select", opts: SN },
      { k: "url", label: "URL del webhook", wide: true, hint: "Zapier, Make, n8n, Slack o un tuo endpoint" },
    ],
    task: [
      { k: "titolo", label: "Task", wide: true },
      { k: "clienteId", label: "Cliente", type: "select", opts: optCliTutti, w2: true },
      { k: "areaId", label: "Area di lavoro", type: "select", opts: optAree },
      { k: "stato", label: "Stato", type: "select", opts: STATI_TASK },
      { k: "priorita", label: "Priorità", type: "select", opts: PRIORITA },
      { k: "stima", label: "Stima (ore)", type: "num", ph: "0,5" },
      { k: "inizio", label: "Inizio", type: "date" },
      { k: "scadenza", label: "Scadenza", type: "date" },
      { k: "ora", label: "Ora (facoltativa)", ph: "15:30", hint: "Con un orario il bot Telegram ti avvisa prima" },
      { k: "ricorrenza", label: "Si ripete", type: "select", opts: RICORRENZE },
      { k: "note", label: "Note", type: "textarea", wide: true },
      { k: "checklist", label: "Passaggi", type: "check", wide: true },
    ],
  };

  const VUOTO = {
    aree: { nome: "" },
    clienti: { nome: "", settore: "", citta: "", stato: "onboarding", cognome: "", padreId: nuovoPadre, archiviato: "no", dataContratto: "", accountAdv: "", fee: 0, budgetMese: 0, margineLordo: 0, resi: 0, aov: 0, aoa: {}, logo: "", note: "", creato: oggi(),
      ragioneSociale: "", piva: "", codiceFiscale: "", indirizzo: "", cap: "", cittaFatt: "", provincia: "", paese: "Italia",
      sdi: "", pec: "", emailAmm: "", referenteAmm: "", termini: "30", metodo: "bonifico" },
    contatti: { clienteId: cliAtt, tipo: "lead", nome: "", azienda: "", ruolo: "", email: "", telefono: "", stato: "nuovo", campagnaId: "", note: "" },
    trattative: { clienteId: cliAtt, titolo: "", contattoId: "", valore: 0, fase: "nuovo", chiusura: piu(30), campagnaId: "", note: "" },
    campagne: { clienteId: cliAtt, nome: "", piattaforma: "meta", obiettivo: "vendite", stato: "test", linea: "", erroreDal: "", ultimaModifica: "", notaModifica: "", notaModificaId: "", modificaVista: "no", funnelId: "", archiviata: "no", inizio: oggi(), fine: piu(30), budgetGiorno: 0, budget: 0, spesa: 0, impression: 0, click: 0, clickLink: 0, lead: 0, conversioni: 0, ricavi: 0, valoreContratto: 0, incassato: 0, frequenza: 0, ...BMK_DEFAULT.vendite },
    vendite: { clienteId: cliAtt, periodo: mese(oggi()), lead: 0, chiamate: 0, discoveryFissate: 0, discoverySvolte: 0, demoFissate: 0, demoSvolte: 0, offerte: 0, chiusure: 0, valoreContratto: 0, incassato: 0, spesa: 0, slot: 0, oreContatto: 0, convPiattaforma: 0 },
    chiusure: { clienteId: cliAtt, data: oggi(), nome: "", azienda: "", email: "", telefono: "", funnel: "form", campagnaId: "", adset: "", annuncio: "", closer: "", valoreContratto: 0, incassato: 0, note: "" },
    eventi: { clienteId: cliAtt, titolo: "", data: oggi(), dataFine: "", ora: "", durata: 1, categoriaId: (data.categorie[0] || {}).id || "", note: "" },
    categorie: { nome: "", colore: "#8B5CF6" },
    modelli: { nome: "", passi: [], creato: oggi() },
    funnel: { clienteId: cliAtt, nome: "", obiettivo: "appuntamento", stato: "test", modello: "", campagnaId: "", spesa: 0, valoreMedio: 0, passi: [], pagina: { titolo: "", sottotitolo: "", punti: "", cta: "Invia la richiesta", webhook: "" }, note: "" },
    annotazioni: { clienteId: cliAtt, data: oggi(), tipo: "budget", campagnaId: "", testo: "", avviaAttesa: false },
    esperimenti: { clienteId: cliAtt, ipotesi: "", variabile: "", metrica: "", inizio: oggi(), fine: "", stato: "corso", risultato: "", conclusione: "" },
    file: { clienteId: cliAtt, titolo: "", categoria: "procedura", tipo: "documento", url: "", tag: "", aggiornato: oggi(), testo: "" },
    periodi: { clienteId: cliAtt, campagnaId: campFasi, tipo: "attiva", inizio: piu(-7), fine: oggi(), spesa: 0, impression: 0, click: 0, clickLink: 0, lead: 0, conversioni: 0, ricavi: 0, valoreContratto: 0, incassato: 0, frequenza: 0, note: "" },
    creativi: { clienteId: cliAtt, campagnaId: "", adsetId: "", tipo: "adset", nome: "", stato: "attiva", stadio: "", spesa: 0, impression: 0, click: 0, clickLink: 0, lead: 0, conversioni: 0, ricavi: 0, valoreContratto: 0, incassato: 0, frequenza: 0, note: "" },
    automazioni: { nome: "", attiva: "si", clienteId: "", trigger: "avviso", filtroAvviso: "", soloCritici: "no", metrica: "cpa", operatore: "gt", valore: 0, giorno: 8, azTask: "si", taskTitolo: "", taskArea: (data.aree[0] || {}).id || "", taskPriorita: "alta", taskGiorni: 2, azWebhook: "no", url: "" },
    task: { clienteId: cliAtt, areaId: (data.aree[0] || {}).id || "", titolo: "", note: "", stato: "dafare", priorita: "media", inizio: oggi(), scadenza: piu(3), stima: 0, ricorrenza: "no", checklist: [], creato: oggi() },
  };

  /* forzaNuovo: rec è un modello precompilato (es. evento in un giorno, sottocliente di X),
     non un record esistente. Così il salvataggio lo aggiunge invece di cercarlo nella lista. */
  const apri = (tipo, rec, forzaNuovo = false) => {
    let d;
    /* Una campagna presa dalle liste calcolate ha totali da fasi, avvisi e data di partenza
       ricalcolata: in modifica si parte sempre dalla scheda salvata. */
    if (tipo === "campagne" && rec && !forzaNuovo) rec = data.campagne.find((x) => x.id === rec.id) || rec;
    if (tipo === "agenzia") d = { ...data.agenzia };
    else if (tipo === "messaggio") d = { consapevolezza: "3", ...((attivo || {}).msg || {}) };
    else if (rec && forzaNuovo) d = { ...VUOTO[tipo], id: uid(), ...rec };
    else if (rec) d = tipo === "campagne" ? { ...rec, ...rec.bmk } : { ...rec };
    else d = { ...VUOTO[tipo], id: uid() };
    if (tipo === "modelli" && rec) {
      d.passi = (rec.passi || []).map(([nome, tp]) => ({ id: uid(), nome, tipo: tp, visite: "", conversioni: "" }));
    }
    if (tipo === "campagne") {
      d.__budget = (d.storicoBudget && d.storicoBudget.length
        ? d.storicoBudget
        : (d.budgetGiorno ? [{ id: uid(), data: d.inizio || oggi(), importo: d.budgetGiorno, nota: "budget iniziale" }] : [])
      ).map((b2) => ({ ...b2, importo: perInput(b2.importo) }));
      d.__fasi = (data.periodi || []).filter((pz) => pz.campagnaId === d.id)
        .map((pz) => ({ id: pz.id, tipo: pz.tipo, inizio: pz.inizio, fine: pz.fine }))
        .sort((a, b) => (a.inizio < b.inizio ? -1 : 1));
    }
    if (tipo === "funnel") {
      d.__campagne = data.campagne.filter((k) => k.funnelId === d.id).map((k) => k.id);
      const p = d.pagina || {};
      d = {
        ...d, pTitolo: p.titolo || "", pSottotitolo: p.sottotitolo || "", pPunti: p.punti || "",
        pCta: p.cta || "Invia la richiesta", pWebhook: p.webhook || "",
        passi: (d.passi || []).map((x) => ({ ...x, visite: perInput(x.visite), conversioni: perInput(x.conversioni) })),
      };
    }
    (NUMERICI[tipo] || []).forEach((k) => { d[k] = perInput(d[k]); });
    setEd({ tipo, nuovo: tipo !== "agenzia" && (!rec || forzaNuovo), draft: d });
  };

  const salva = () => {
    const { tipo, draft, nuovo } = ed;
    let d = { ...draft };
    (NUMERICI[tipo] || []).forEach((k) => { d[k] = parseNum(d[k]); });
    if (tipo === "agenzia") { setData((x) => ({ ...x, agenzia: d })); setEd(null); return; }
    if (tipo === "messaggio") {
      upd("clienti", (l) => l.map((z) => (z.id === cliAtt ? { ...z, msg: d } : z)));
      setEd(null); return;
    }
    if (tipo === "campagne") {
      if (!d.clienteId) {
        chiedi({ titolo: "Manca il cliente", testo: "Scegli a quale cliente appartiene la campagna prima di salvarla.", etichetta: "Ho capito", azione: () => {} });
        return;
      }
      const { roasMin = 0, cpaMax = 0, cplMax = 0, ctrMin = 0, freqMax = 0, cpmMax = 0, __fasi = [], __budget = [], d: _d, al: _al, daFasi: _f, inizioMeta: _im, ...resto } = d;
      const storico = __budget.filter((b2) => b2.data).map((b2) => ({ ...b2, importo: parseNumIntl(b2.importo) }))
        .sort((a, b) => (a.data < b.data ? -1 : 1));
      /* Se hai cambiato a mano il budget giornaliero, il cambio entra nello storico con la data di oggi:
         prima veniva sovrascritto in silenzio dall'ultima voce dello storico. */
      const originale = data.campagne.find((x) => x.id === resto.id);
      const scritto = resto.budgetGiorno || 0;
      const toccato = originale ? scritto !== (originale.budgetGiorno || 0) : !storico.length;
      if (toccato && scritto > 0 && (!storico.length || storico[storico.length - 1].importo !== scritto)) {
        const dataVoce = originale ? oggi() : (resto.inizio || oggi());
        const i = storico.findIndex((b2) => b2.data === dataVoce);
        if (i > -1) storico[i] = { ...storico[i], importo: scritto };
        else storico.push({ id: uid(), data: dataVoce, importo: scritto, nota: originale ? "modificato dalla scheda" : "budget iniziale" });
        storico.sort((a, b) => (a.data < b.data ? -1 : 1));
      }
      resto.storicoBudget = storico;
      if (storico.length) resto.budgetGiorno = storico[storico.length - 1].importo;
      const ultimo = [...__fasi].sort((a, b) => (a.fine < b.fine ? -1 : 1)).slice(-1)[0];
      /* Se la data di fine l'hai scritta tu vale quella: è la chiusura reale della campagna,
         non il bordo dell'ultima finestra compilata. */
      /* Traccia del fermo: quando entra in errore segno il giorno, quando ne esce
         scrivo nel Diario quanto è durato. Così resta la storia, non solo lo stato. */
      const prima = data.campagne.find((x) => x.id === resto.id);
      const eraErrore = prima && prima.stato === "errore";
      if (!eraErrore && resto.stato === "errore" && !resto.erroreDal) resto.erroreDal = oggi();
      if (eraErrore && resto.stato !== "errore") {
        const dal = prima.erroreDal || oggi();
        const g = giorniTra(dal, oggi());
        upd("annotazioni", (l) => [{
          id: uid(), clienteId: resto.clienteId, campagnaId: resto.id, data: oggi(), tipo: "esterno",
          testo: `${causaFermo(prima).l} risolto su "${resto.nome}". Campagna ferma dal ${dataItA(dal)} al ${dataItA(oggi())}: ${int(g)} ${g === 1 ? "giorno" : "giorni"} senza erogazione. Stato riportato su "${et(STATI_CAMPAGNA, resto.stato)}".`,
        }, ...l]);
        resto.erroreDal = "";
        resto.causaErrore = "";
      }
      /* I totali salvati seguono le finestre tenute: se ne cancelli una, il totale
         in archivio scende con lei invece di restare il vecchio numero. */
      const tenuti = __fasi.map((x) => x.id);
      const CH = ["spesa", "impression", "click", "clickLink", "lead", "conversioni", "ricavi"];
      const sf = { n: 0, fq: 0 };
      CH.forEach((c) => { sf[c] = 0; });
      (data.periodi || []).filter((pz) => pz.campagnaId === resto.id && tenuti.includes(pz.id)).forEach((pz) => {
        sf.n += 1;
        CH.forEach((c) => { sf[c] += pz[c] || 0; });
        sf.fq += (pz.frequenza || 0) * (pz.impression || 0);
      });
      if (sf.n && sf.spesa + sf.impression + sf.click + sf.conversioni + sf.ricavi > 0) {
        /* Conversioni e ricavi sono tuoi se li hai scritti diversi dalla somma:
           in una campagna a contatti il lead non è una conversione, e il numero
           che conta — appuntamenti, vendite — lo sai solo tu. */
        const mio = (resto.conversioni || 0) !== sf.conversioni || (resto.ricavi || 0) !== sf.ricavi;
        CH.filter((c) => !(mio && (c === "conversioni" || c === "ricavi")))
          .forEach((c) => { resto[c] = sf[c]; });
        resto.esitoManuale = mio ? "si" : "";
        if (sf.impression) resto.frequenza = Math.round((sf.fq / sf.impression) * 100) / 100;
      }
      const fineScritta = (resto.fine || "").trim();
      d = { ...resto, bmk: { roasMin, cpaMax, cplMax, ctrMin, freqMax, cpmMax },
        fine: fineScritta || (ultimo ? ultimo.fine : (resto.inizio ? sommaGiorni(resto.inizio, 30) : oggi())) };
      /* Le finestre seguono la scheda: aggiorno le date, creo le nuove, tolgo quelle rimosse */
      const idsTenuti = __fasi.map((x) => x.id);
      upd("periodi", (l) => [
        ...l.filter((pz) => pz.campagnaId !== d.id || idsTenuti.includes(pz.id))
          .map((pz) => {
            const f2 = __fasi.find((x) => x.id === pz.id);
            return f2 && pz.campagnaId === d.id ? { ...pz, tipo: f2.tipo, inizio: f2.inizio, fine: f2.fine } : pz;
          }),
        ...__fasi.filter((x) => !l.some((pz) => pz.id === x.id)).map((x) => ({
          id: x.id, clienteId: d.clienteId, campagnaId: d.id, tipo: x.tipo, inizio: x.inizio, fine: x.fine,
          spesa: 0, impression: 0, click: 0, conversioni: 0, ricavi: 0, valoreContratto: 0, incassato: 0, frequenza: 0, note: "",
        })),
      ]);
    }
    if (tipo === "task") d.checklist = (d.checklist || []).filter((r) => r.t.trim());
    if (tipo === "creativi" && d.campagnaId) {
      const k = data.campagne.find((x) => x.id === d.campagnaId);
      if (k) d.clienteId = k.clienteId;
    }
    if (tipo === "funnel") {
      const { pTitolo, pSottotitolo, pPunti, pCta, pWebhook, __campagne = [], ...resto } = d;
      d = {
        ...resto,
        pagina: { titolo: pTitolo, sottotitolo: pSottotitolo, punti: pPunti, cta: pCta, webhook: pWebhook },
        passi: (d.passi || []).filter((x) => (x.nome || "").trim())
          .map((x) => ({ ...x, visite: parseNum(x.visite), conversioni: parseNum(x.conversioni) })),
      };
      /* Le campagne spuntate puntano qui, quelle tolte si liberano */
      upd("campagne", (l) => l.map((k) => {
        if (__campagne.includes(k.id)) return { ...k, funnelId: d.id };
        if (k.funnelId === d.id) return { ...k, funnelId: "" };
        return k;
      }));
    }
    if (tipo === "annotazioni") {
      /* Una nota di diario su una campagna può far partire i tre giorni di attesa */
      const { avviaAttesa, ...resto } = d;
      if (avviaAttesa === true && d.campagnaId) {
        upd("campagne", (l) => l.map((x) => (x.id === d.campagnaId
          ? { ...x, ultimaModifica: d.data || oggi(), notaModifica: d.testo || "", notaModificaId: d.id } : x)));
      }
      d = resto;
    }
    if (tipo === "modelli") {
      d = {
        ...d, creato: d.creato || oggi(),
        passi: (d.passi || []).filter((x) => (x.nome || "").trim()).map((x) => [x.nome, x.tipo]),
      };
    }
    upd(tipo, (l) => (nuovo ? [d, ...l] : l.map((x) => (x.id === d.id ? d : x))));
    setNuovoPadre(""); setEd(null);
  };

  const elimina = () => {
    const { tipo, draft } = ed;
    if (tipo === "annotazioni" && data.campagne.some((k) => k.notaModificaId === draft.id)) {
      /* Cancellare la nota cancella anche l'attesa che aveva avviato: un riferimento
         a una nota che non esiste più terrebbe la campagna sospesa senza motivo. */
      upd("campagne", (l) => l.map((k) => (k.notaModificaId === draft.id
        ? { ...k, ultimaModifica: "", notaModifica: "", notaModificaId: "" } : k)));
    }
    if (tipo === "clienti") {
      setEd(null);
      eliminaCliente(draft);
      return;
    }
    if (tipo === "aree" && data.task.some((t) => t.areaId === draft.id)) {
      setEd(null);
      chiedi({
        titolo: "Eliminare l'area?",
        testo: "Ci sono task in quest'area: resteranno senza area assegnata.",
        etichetta: "Elimina", pericolo: true,
        azione: () => upd("aree", (l) => l.filter((x) => x.id !== draft.id)),
      });
      return;
    }
    upd(tipo, (l) => l.filter((x) => x.id !== draft.id));
    setEd(null);
  };

  const muovi = (t, dir) => {
    const i = FASI.findIndex((f) => f.v === t.fase);
    const next = FASI[Math.min(FASI.length - 1, Math.max(0, i + dir))];
    upd("trattative", (l) => l.map((x) => (x.id === t.id ? { ...x, fase: next.v } : x)));
  };

  const cambiaStatoTask = (t, stato) => {
    const chiuso = stato === "fatto" && t.stato !== "fatto";
    const g = (RICORRENZE.find((r) => r.v === t.ricorrenza) || {}).g || 0;
    setData((d) => {
      let lista = d.task.map((x) => (x.id === t.id ? { ...x, stato } : x));
      if (chiuso && g) {
        lista = [{
          ...t, id: uid(), stato: "dafare",
          inizio: sommaGiorni(t.inizio || oggi(), g),
          scadenza: sommaGiorni(t.scadenza || oggi(), g),
          checklist: (t.checklist || []).map((r) => ({ ...r, fatto: false })), creato: oggi(),
        }, ...lista];
      }
      return { ...d, task: lista };
    });
  };
  const spostaTask = (t, dir) => {
    const i = STATI_TASK.findIndex((s) => s.v === t.stato);
    cambiaStatoTask(t, STATI_TASK[Math.min(STATI_TASK.length - 1, Math.max(0, i + dir))].v);
  };

  const reset = () => {
    chiedi({
      titolo: "Ripartire dai dati di esempio?",
      testo: "Tutto quello che hai inserito viene sostituito con l'archivio dimostrativo. Se non hai un backup recente, annulla e scaricalo prima.",
      etichetta: "Sostituisci tutto", pericolo: true,
      azione: () => { sbloccaArchivio(); setCliAtt(""); setVista("cruscotto"); setData(normalizza(SEED())); },
    });
  };
  const cambiaAccount = (id) => {
    setCliAtt(id); setVista("cruscotto"); setQ(""); setConsolidato(true);
    /* Se entri in un sottocliente, la tendina del suo madre resta aperta */
    const z = data.clienti.find((x) => x.id === id);
    if (z && z.padreId) setApertiCli((x) => ({ ...x, [z.padreId]: true }));
  };

  const NAV_AG = [
    { id: "cruscotto", l: t("nav.cruscotto"), i: "◱" },
    { id: "clienti", l: t("nav.clienti"), i: "▣", n: attivi.filter((z) => !z.padreId).length },
    { id: "sottoclienti", l: t("nav.sottoclienti"), i: "▤", n: attivi.filter((z) => z.padreId).length },
    { id: "advertising", l: t("nav.advertising"), i: "◈", n: m.campagne.length },
    { id: "mesi", l: t("nav.mesi"), i: "▤" },
    { id: "semaforo", l: t("nav.semaforo"), i: "◉" },
    { id: "pixel", l: t("nav.pixel"), i: "◈" },
    { id: "calendario", l: t("nav.calendario"), i: "▦", n: data.eventi.length },
    { id: "avvisi", l: t("nav.avvisi"), i: "!", n: m.avvisi.length, al: m.avvisi.length > 0 },
    { id: "grafici", l: t("nav.grafici"), i: "◔" },
    { id: "vendite", l: t("nav.vendite"), i: "⇲", n: data.vendite.length },
    { id: "chiusure", l: t("nav.chiusure"), i: "€", n: data.chiusure.length },
    { id: "pianificatore", l: t("nav.pianificatore"), i: "⌁" },
    { id: "funnel", l: t("nav.funnel"), i: "▷", n: data.funnel.length },
    { id: "diario", l: t("nav.diario"), i: "✎", n: data.annotazioni.length },
    { id: "file", l: t("nav.file"), i: "▦", n: data.file.length },
    { id: "economia", l: t("nav.economia"), i: "◑" },
    { id: "ai", l: t("nav.ai"), i: "✳" },
    { id: "task", l: t("nav.task"), i: "✓", n: data.task.filter((t) => t.stato !== "fatto").length },
    { id: "automazioni", l: t("nav.automazioni"), i: "⚡", n: data.automazioni.filter((a) => a.attiva === "si").length },
    { id: "importa", l: t("nav.importa"), i: "⇪" },
    { id: "backup", l: t("nav.backup"), i: "⤓" },
    { id: "academy", l: t("nav.academy"), i: "🎓" },
    { id: "guida", l: t("nav.guida"), i: "?" },
    { id: "impostazioni", l: t("nav.impostazioni"), i: "⚙" },
  ];
  const NAV_CLI = [
    { id: "cruscotto", l: t("nav.cruscotto"), i: "◱" },
    { id: "contatti", l: t("nav.contatti"), i: "◉", n: S.contatti.length },
    { id: "pipeline", l: t("nav.pipeline"), i: "≡", n: S.trattative.filter((t) => !["vinto", "perso"].includes(t.fase)).length },
    { id: "messaggio", l: t("nav.messaggio"), i: "✍", n: (attivo && attivo.msg && attivo.msg.mercato) ? undefined : "!", al: !(attivo && attivo.msg && attivo.msg.mercato) },
    { id: "campagne", l: t("nav.campagne"), i: "◈", n: S.campagne.length },
    { id: "winner", l: t("nav.winner"), i: "★" },
    { id: "mesi", l: t("nav.mesi"), i: "▤" },
    { id: "semaforo", l: t("nav.semaforo"), i: "◉" },
    { id: "pixel", l: t("nav.pixel"), i: "◈" },
    { id: "calendario", l: t("nav.calendario"), i: "▦", n: data.eventi.length },
    { id: "avvisi", l: t("nav.avvisi"), i: "!", n: S.avvisi.length, al: S.avvisi.length > 0 },
    { id: "grafici", l: t("nav.grafici"), i: "◔" },
    { id: "vendite", l: t("nav.vendite"), i: "⇲", n: S.vendite.length },
    { id: "chiusure", l: t("nav.chiusure"), i: "€", n: S.chiusure.length },
    { id: "pianificatore", l: t("nav.pianificatore"), i: "⌁" },
    { id: "funnel", l: t("nav.funnel"), i: "▷", n: S.funnel.length },
    { id: "diario", l: t("nav.diario"), i: "✎", n: S.annotazioni.length },
    { id: "file", l: t("nav.file"), i: "▦", n: S.file.length },
    { id: "report", l: t("nav.report"), i: "▥" },
    { id: "ai", l: t("nav.ai"), i: "✳" },
    { id: "task", l: t("nav.task"), i: "✓", n: S.task.filter((t) => t.stato !== "fatto").length },
    { id: "scheda", l: t("nav.scheda"), i: "▤" },
    { id: "guida", l: t("nav.guida"), i: "?" },
  ];
  const NAV = attivo ? NAV_CLI : NAV_AG;

  const titoli = attivo ? {
    cruscotto: [attivo.nome, "Tutto quello che riguarda questo cliente"],
    contatti: ["Contatti", "Lead generati e referenti dell'azienda"],
    pipeline: ["Pipeline", "Le trattative aperte di questo cliente"],
    messaggio: ["Messaggio", "Chi è il mercato, cosa gli dici, e i quattro annunci"],
    campagne: ["Advertising", "Campagne, ritmo di spesa, ad set e annunci"],
    winner: ["Winner ads", "La classifica degli annunci sui benchmark e le analisi del Claude del cliente"],
    mesi: ["Riepilogo mensile", "Spesa, vendite, contratti e incassato mese per mese"],
    semaforo: ["Sistema a semaforo", "Due indicatori, tre risposte: spegni, aspetta, scala"],
    pixel: ["Pixel facile", "Genera il codice degli eventi standard di Meta"],
    calendario: ["Calendario", "Scadenze, avvii, finestre e chiusure su un unico piano"],
    avvisi: ["Avvisi", "Soglie superate, causa probabile e cosa fare"],
    grafici: ["Grafici", "Scegli cosa misurare e come rappresentarlo"],
    vendite: ["Ciclo di vendita", "Chiamate, discovery, demo e chiusure mese per mese"],
    chiusure: ["Chiusure", "Chi hai chiuso, per quanto, e da quale annuncio è arrivato"],
    pianificatore: ["Pianificatore", "Parti dall'obiettivo e risali a lead e budget"],
    funnel: ["Funnel", "La sequenza passo per passo, e dove si perdono le persone"],
    diario: ["Diario", "Cosa hai cambiato, e cosa hai imparato"],
    file: ["File", "Strategie, procedure e collegamenti ai documenti"],
    report: ["Report", "Il mese da consegnare al cliente"],
    ai: ["Consulente", "Chiedi, con i tuoi numeri già sotto gli occhi"],
    task: ["Task", "Il lavoro da fare su questo cliente"],
    scheda: ["Scheda cliente", "Dati, logo, compenso e vincoli"],
    guida: ["Guida", "Come si usa, in che ordine, e cosa non può fare"],
  } : {
    cruscotto: ["Panoramica agenzia", "Come stanno andando tutti gli account"],
    sottoclienti: ["Sottoclienti", "Gli account secondari, raggruppati per cliente principale"],
    clienti: ["Clienti", "Entra in un account per avere CRM e marketing dedicati"],
    advertising: ["Advertising", "I totali di ogni cliente: campagne, ad set e annunci sono dentro il cliente"],
    avvisi: ["Avvisi", "Tutte le soglie superate, su tutti i clienti"],
    mesi: ["Riepilogo mensile", "Spesa, vendite, contratti e incassato di tutto il portafoglio"],
    semaforo: ["Sistema a semaforo", "Due indicatori, tre risposte: spegni, aspetta, scala"],
    pixel: ["Pixel facile", "Genera il codice degli eventi standard di Meta"],
    calendario: ["Calendario", "Scadenze, avvii, finestre e chiusure di tutti i clienti"],
    grafici: ["Grafici", "Riepiloghi su tutto il portafoglio"],
    vendite: ["Ciclo di vendita", "I totali commerciali di ogni cliente"],
    chiusure: ["Chiusure", "Clienti chiusi, contratti e incassato di ogni cliente"],
    pianificatore: ["Pianificatore", "Dall'obiettivo di chiusure al budget necessario"],
    funnel: ["Funnel", "Tutti i percorsi di conversione del portafoglio"],
    diario: ["Diario", "Decisioni prese ed esperimenti, su tutti i clienti"],
    file: ["File", "Procedure dell'agenzia e documenti dei clienti"],
    economia: ["Economia", "Quanto rende ogni cliente, e quanto ti costa"],
    ai: ["Consulente", "Chiedi, con i numeri del portafoglio già sotto gli occhi"],
    task: ["Task", "Tutto il lavoro, per area e per scadenza"],
    automazioni: ["Automazioni", "Regole che lavorano da sole: innesco e azione"],
    backup: ["Backup", "Scarica una copia, o ripristinane una"],
    importa: ["Importa dati", "Dai file esportati dalle piattaforme, senza riscrivere nulla"],
    academy: ["Academy", "I corsi dell'Accademia dell'Advertising, a portata di clic"],
    guida: ["Guida", "Come si usa, in che ordine, e cosa non può fare"],
    impostazioni: ["Impostazioni", "Identità dell'agenzia e aree di lavoro"],
  };

  const cerca = (s) => (s || "").toLowerCase().includes(q.toLowerCase());
  const bottoneAggiungi = attivo
    ? { contatti: "contatti", pipeline: "trattative", campagne: "campagne", task: "task", vendite: "vendite", chiusure: "chiusure", funnel: "funnel", file: "file" }[vista]
    : { clienti: "clienti", advertising: "campagne", task: "task", vendite: "vendite", chiusure: "chiusure", automazioni: "automazioni", funnel: "funnel", file: "file" }[vista];

  const taskBase = (attivo ? S.task : data.task).filter((t) =>
    (!filtri.area || t.areaId === filtri.area) &&
    (!filtri.prio || t.priorita === filtri.prio) &&
    (attivo || !filtri.cli || t.clienteId === filtri.cli) &&
    (filtri.fatte || t.stato !== "fatto")
  );
  const ordinaTask = (l) => [...l].sort((a, b) => {
    const sa = a.scadenza || "9999-99-99", sb = b.scadenza || "9999-99-99";
    if (sa !== sb) return sa < sb ? -1 : 1;
    return prioR(a.priorita) - prioR(b.priorita);
  });
  const gruppiTask = () => {
    const l = ordinaTask(taskBase.filter((t) => t.stato !== "fatto"));
    const fine = piu(7);
    const g = [
      { t: "In ritardo", items: l.filter((x) => x.scadenza && x.scadenza < oggi()) },
      { t: "Oggi", items: l.filter((x) => x.scadenza === oggi()) },
      { t: "Prossimi sette giorni", items: l.filter((x) => x.scadenza > oggi() && x.scadenza <= fine) },
      { t: "Più avanti", items: l.filter((x) => x.scadenza > fine) },
      { t: "Senza scadenza", items: l.filter((x) => !x.scadenza) },
    ];
    if (filtri.fatte) g.push({ t: "Fatte", items: ordinaTask(taskBase.filter((x) => x.stato === "fatto")) });
    return g.filter((x) => x.items.length);
  };

  const Barra = ({ perf }) => (
    <div className="track">
      <div className="fill" data-sotto={perf < 1 ? "1" : "0"} style={{ width: `${Math.min(100, (perf / 1.5) * 100)}%` }} />
      <div className="mark" style={{ left: "66.6%" }} />
    </div>
  );

  const TestaCampagne = () => (
    <thead><tr>
      <th /><th>Campagna</th><th>Stato</th><th className="r">Speso</th><th className="r">CTR link</th><th className="r">CPC link</th>
      <th className="r">CPM</th><th className="r">Lead</th><th className="r">CPL</th>
      <th className="r">Conv.</th><th className="r">CPA</th><th className="r">ROAS</th><th className="r">Avvisi</th>
    </tr></thead>
  );

  /* Riga di campagna: nella valuta del suo cliente (un sottocliente può averne una diversa dalla madre) */
  const RigaCampagna = ({ k }) => conValuta(valutaCliente(data, k.clienteId), () => (
    <tr className="clic" onClick={() => apri("campagne", data.campagne.find((x) => x.id === k.id))}>
      <td style={{ paddingRight: 0, width: 22 }}>
        {(() => {
          const sg = semaforo(data, k);
          return <span title={`${sg.etichetta} — ${sg.perche}`}
            style={{ width: 9, height: 9, borderRadius: "50%", background: sg.colore, display: "inline-block" }} />;
        })()}
      </td>
      <td>
        <div className="riga" style={{ gap: 6 }}><b>{k.nome}</b><Campanella k={k} size={12} /></div>
        <div className="sub">
          {et(PIATTAFORME, k.piattaforma)} · {et(OBIETTIVI, k.obiettivo)} · {faseCampagna(k).l}
        </div>
        {(() => {
          const pr = valutazionePronta(k);
          if (pr) return <div className="sub" style={{ color: "#C99A3A" }}>valutazione pronta dal {dataIt(pr.scaduta)}</div>;
          const w = inValutazione(k);
          return w ? <div className="sub" style={{ color: "#C99A3A" }}>in valutazione dopo modifica · fino al {dataIt(w.fino)}</div> : null;
        })()}
      </td>
      <td><span className="tag" data-k={k.stato}>{etStato(k)}</span>{k.accountBloccato ? <div className="acc-bloccato" title={k.stato === "errore" ? "La campagna era accesa: l'account l'ha fermata" : "La campagna era già spenta: l'account non la ferma, ma se la riaccendi non erogherà"}>⚠ account {k.accountBloccato}{k.stato !== "errore" ? " · campagna già spenta" : ""}</div> : null}</td>
      <td className="r">{eur(k.d.sp)}<div className="sub mono">di {eur(k.budget)}</div></td>
      <td className={"r " + (k.bmk.ctrMin && (k.d.ctrLink || k.d.ctr) < k.bmk.ctrMin ? "neg" : "")}>{pctv(k.d.ctrLink || k.d.ctr)}</td>
      <td className="r">{eur(k.d.cpcLink || k.d.cpc)}</td>
      <td className={"r " + (k.bmk.cpmMax && k.d.cpm > k.bmk.cpmMax ? "neg" : "")}>{eur(k.d.cpm)}</td>
      <td className="r">{k.d.ld ? int(k.d.ld) : "—"}</td>
      <td className={"r " + (tettoCpl(k.bmk) && k.d.ld && k.d.cpl > tettoCpl(k.bmk) ? "neg" : "")}>{k.d.ld ? eur(k.d.cpl) : "—"}</td>
      <td className="r">
        {int(k.d.cv)}
        {k.esitoManuale === "si" && <div className="sub">a mano</div>}
        {k.d.cv > 0 && k.d.cv === k.d.ld && k.esitoManuale !== "si" && <div className="sub">= lead</div>}
      </td>
      <td className={"r " + (k.bmk.cpaMax && k.d.cv && k.d.cpa > k.bmk.cpaMax ? "neg" : "")}>{k.d.cv ? eur(k.d.cpa) : "—"}</td>
      <td className={"r " + (k.d.rc ? (k.bmk.roasMin && k.d.roas < k.bmk.roasMin ? "neg" : "pos") : "")}>{k.d.rc ? `${fmt(k.d.roas)}×` : "—"}</td>
      <td className="r">{veri(k.al).length ? <span className="tag" data-k={k.al.some((a) => a.gravita === "critica") ? "late" : ""}>{veri(k.al).length}</span> : <span className="sub">ok</span>}</td>
    </tr>
  ));

  /* Task arrivate da un software esterno (Boschetto): da dove vengono, con il link se c'è */
  const badgeEsterno = (t) => (!t.esterno ? null : t.esterno.url
    ? <a className="pill" href={t.esterno.url} target="_blank" rel="noopener noreferrer" onClick={(e) => e.stopPropagation()} title={`Apri in ${t.esterno.fonte}${t.esterno.cliente ? ` (${t.esterno.cliente})` : ""}`}>da {t.esterno.fonte} ↗</a>
    : <span className="pill" title={t.esterno.cliente ? `In ${t.esterno.fonte}: ${t.esterno.cliente}` : ""}>da {t.esterno.fonte}</span>);
  const CartaTask = ({ t }) => {
    const done = (t.checklist || []).filter((r) => r.fatto).length, tot = (t.checklist || []).length;
    const late = t.scadenza && t.scadenza < oggi() && t.stato !== "fatto";
    const i = STATI_TASK.findIndex((s) => s.v === t.stato);
    return (
      <article className="tk" data-p={t.priorita} data-scaduta={late ? "1" : "0"}>
        <b>{t.titolo}</b>
        <div className="meta">
          <span className="tag" data-k={t.priorita}>{et(PRIORITA, t.priorita)}</span>
          <span className="pill" data-late={late ? "1" : "0"}>{t.scadenza ? dataIt(t.scadenza) : "senza data"}</span>
          {!attivo && <span className="pill">{t.clienteId ? nomeCli(t.clienteId) : "Agenzia"}</span>}
          <span className="pill">{nomeArea(t.areaId)}</span>
          {t.ricorrenza !== "no" && <span className="pill">↻</span>}
          {badgeEsterno(t)}
        </div>
        {tot > 0 && (
          <>
            <div className="barra"><i style={{ width: `${(done / tot) * 100}%` }} /></div>
            <div className="sub" style={{ marginTop: 4 }}>{done} di {tot} passaggi</div>
          </>
        )}
        <div className="row">
          {t.stima > 0 && <span className="sub mono">{fmt(t.stima, 1)} h</span>}
          <span style={{ marginLeft: "auto", display: "flex", gap: 4 }}>
            <button className="move" onClick={() => spostaTask(t, -1)} disabled={i === 0} aria-label="Stato precedente">‹</button>
            <button className="move" onClick={() => spostaTask(t, 1)} disabled={i === STATI_TASK.length - 1} aria-label="Stato successivo">›</button>
            <button className="move" onClick={() => apri("task", t)} aria-label="Apri">✎</button>
          </span>
        </div>
      </article>
    );
  };

  const RigaTask = ({ t }) => {
    const late = t.scadenza && t.scadenza < oggi() && t.stato !== "fatto";
    const done = (t.checklist || []).filter((r) => r.fatto).length, tot = (t.checklist || []).length;
    return (
      <div className="tkrow" data-done={t.stato === "fatto" ? "1" : "0"}>
        <input type="checkbox" checked={t.stato === "fatto"} aria-label="Fatto"
          onChange={(e) => cambiaStatoTask(t, e.target.checked ? "fatto" : "dafare")} />
        <div className="t clic" onClick={() => apri("task", t)}>
          <b>{t.titolo}</b>
          <div className="meta" style={{ marginTop: 5 }}>
            <span className="tag" data-k={t.priorita}>{et(PRIORITA, t.priorita)}</span>
            <span className="pill">{nomeArea(t.areaId)}</span>
            {!attivo && <span className="pill">{t.clienteId ? nomeCli(t.clienteId) : "Agenzia"}</span>}
            <span className="pill">{et(STATI_TASK, t.stato)}</span>
            {tot > 0 && <span className="pill">{done}/{tot}</span>}
            {t.ricorrenza !== "no" && <span className="pill">↻</span>}
            {badgeEsterno(t)}
          </div>
        </div>
        <span className="pill" data-late={late ? "1" : "0"}>{t.scadenza ? dataIt(t.scadenza) : "—"}</span>
      </div>
    );
  };

  const SezioneGrafici = () => {
    const G = costruisci();
    const torta = graf.tipo === "torta" && !G.rap;
    const baseCamp = attivo ? S.campagne : m.campagne;
    const perQuota = new Map();
    baseCamp.forEach((k) => {
      const key = attivo ? et(PIATTAFORME, k.piattaforma) : nomeCli(k.clienteId);
      perQuota.set(key, (perQuota.get(key) || 0) + k.d.sp);
    });
    const quota = [...perQuota.entries()].map(([k, v]) => ({ k, v })).sort((a, b) => b.v - a.v);
    const tot = sommaVendite(venditeVista);
    const imbuto = FASI_VENDITA.map((f) => ({ k: f.l, v: tot[f.k] || 0 }));
    const mf = new Map();
    chiusureVista.forEach((c) => mf.set(et(FUNNEL_PROV, c.funnel), (mf.get(et(FUNNEL_PROV, c.funnel)) || 0) + (c.valoreContratto || 0)));
    const perFunnel = [...mf.entries()].map(([k, v]) => ({ k, v })).sort((a, b) => b.v - a.v);

    return conValuta(valutaComune(attivo ? ambito : baseCamp.map((k) => k.clienteId)), () => (
      <>
        <section className="card">
          <header>
            <div className="ctrl">
              <select value={graf.ambito} onChange={(e) => cambiaGraf({ ambito: e.target.value })} aria-label="Cosa misurare">
                <option value="campagne">Campagne pubblicitarie</option>
                <option value="vendite">Ciclo di vendita</option>
                <option value="chiusure">Chiusure e provenienza</option>
              </select>
              <select value={graf.gruppo} onChange={(e) => cambiaGraf({ gruppo: e.target.value })} aria-label="Come raggruppare">
                {gruppiDisponibili.map((g) => <option key={g.v} value={g.v}>{g.l}</option>)}
              </select>
              <select value={graf.metrica} onChange={(e) => cambiaGraf({ metrica: e.target.value })} aria-label="Metrica">
                {metricheDisponibili.map((x) => <option key={x.v} value={x.v}>{x.l}</option>)}
              </select>
              <button className="btn ghost mini" data-on={graf.tipo === "barre" ? "1" : "0"} onClick={() => cambiaGraf({ tipo: "barre" })}>Barre</button>
              <button className="btn ghost mini" data-on={graf.tipo === "torta" ? "1" : "0"} onClick={() => cambiaGraf({ tipo: "torta" })}>Torta</button>
            </div>
          </header>
          <div className="gruppo">{G.titolo}</div>
          {torta ? <Torta dati={G.dati} fmtv={fissaValuta(G.fmtv)} /> : <Barre dati={G.dati} fmtv={fissaValuta(G.fmtv)} />}
          {graf.tipo === "torta" && G.rap && (
            <div className="nota">
              Questa metrica è un rapporto, non una quantità: le fette non possono sommare al totale, quindi la mostro a barre.
              Per una torta scegli una misura che si somma, come spesa, ricavi o chiusure.
            </div>
          )}
        </section>

        <div className="two">
          <section className="card">
            <header><h2>Dove va la spesa</h2><span className="sp eyebrow">{attivo ? "per piattaforma" : "per cliente"}</span></header>
            <Torta dati={quota} fmtv={fissaValuta(eur)} />
          </section>
          <section className="card">
            <header><h2>Percorso commerciale</h2><span className="sp eyebrow">totale del periodo caricato</span></header>
            {venditeVista.length
              ? <Barre dati={imbuto} fmtv={int} />
              : <div className="empty"><b>Nessun dato commerciale</b>Carica un mese dalla sezione Ciclo di vendita.</div>}
          </section>
        </div>

        <section className="card">
          <header><h2>Da dove arriva il fatturato chiuso</h2><span className="sp eyebrow">valore contratto per funnel</span></header>
          {chiusureVista.length
            ? <Barre dati={perFunnel} fmtv={fissaValuta(eur)} />
            : <div className="empty"><b>Nessuna chiusura registrata</b>Le schede della sezione Chiusure riempiono questo riquadro.</div>}
        </section>
      </>
    ));
  };

  /* ---------- sistema a semaforo ---------- */
  const SezioneSemaforo = () => {
    const cid = attivo ? attivo.id : semCliente;
    const disponibili = (!cid ? [] : m.campagne.filter((k) => (attivo ? ambito.includes(k.clienteId) : k.clienteId === cid)))
      .filter((k) => k.archiviata !== "si")
      .sort((a, b) => (b.inizio || "").localeCompare(a.inizio || ""));
    const scelte = semCamp.length ? disponibili.filter((k) => semCamp.includes(k.id)) : disponibili;
    const esiti = scelte.map((k) => ({ k, s: semaforo(data, k) }));
    const conta = (v) => esiti.filter((x) => x.s.segnale === v).length;
    const segna = (v) => (SEGNALI.find((x) => x.v === v) || {});

    const Pallino = ({ v, size = 10 }) => (
      <span aria-label={segna(v).l} title={segna(v).l}
        style={{ width: size, height: size, borderRadius: "50%", background: segna(v).c, flex: "none", display: "inline-block" }} />
    );

    /* Fuori dal cliente si guarda un cliente scelto: tutta la sezione parla nella sua valuta */
    return conValuta(cid ? valutaCliente(data, cid) : VALUTA, () => (
      <>
        <section className="card">
          <header>
            <h2>Cosa analizzare</h2>
            <span className="sp eyebrow">{esiti.length} campagne sotto esame</span>
            {disponibili.length > 0 && (
              <>
                <button className="btn ghost mini" onClick={() => setSemCamp([])}>Tutte</button>
                <button className="btn ghost mini"
                  onClick={() => setSemCamp(disponibili.filter((k) => ["attiva", "test"].includes(k.stato)).map((k) => k.id))}>
                  Solo le attive
                </button>
              </>
            )}
          </header>

          <div className="pannello-filtri">
            {!attivo && (
              <div className="f">
                <label htmlFor="sem-cli">Account da analizzare</label>
                <select id="sem-cli" value={semCliente} onChange={(e) => { setSemCliente(e.target.value); setSemCamp([]); }}>
                  <option value="">— scegli un cliente —</option>
                  {ordinaGerarchia().map(({ z, liv }) => (
                    <option key={z.id} value={z.id}>{liv ? `${"\u00A0".repeat(liv * 4)}↳ ` : ""}{z.nome}</option>
                  ))}
                </select>
              </div>
            )}
            <div className="f">
              <label>In sintesi</label>
              <div className="riga" style={{ gap: 16, paddingTop: 7 }}>
                {SEGNALI.map((g) => (
                  <span className="riga" key={g.v} style={{ gap: 7, fontSize: 12.8 }}>
                    <Pallino v={g.v} /> <b className="mono">{conta(g.v)}</b> <span className="sub">{g.l.toLowerCase()}</span>
                  </span>
                ))}
              </div>
            </div>
          </div>

          {!cid ? (
            <div className="empty"><b>Scegli un account</b>Il semaforo si applica alle campagne di un cliente alla volta.</div>
          ) : !disponibili.length ? (
            <div className="empty"><b>Nessuna campagna</b>Questo account non ha campagne da analizzare.</div>
          ) : (
            <>
              <div className="sceltarigaintest" style={{ margin: "12px 16px 0" }}>
                <span /><span>Campagna</span><span>Fase</span><span>Segnale</span>
              </div>
              <div className="sceltalista" style={{ maxHeight: 280, padding: "0 16px 14px" }}>
                {disponibili.map((k) => {
                  const on = !semCamp.length || semCamp.includes(k.id);
                  const s = semaforo(data, k);
                  return (
                    <label className="sceltariga" key={k.id} data-on={on ? "1" : "0"}>
                      <input type="checkbox" checked={on}
                        onChange={(e) => {
                          const base = semCamp.length ? semCamp : disponibili.map((x) => x.id);
                          setSemCamp(e.target.checked ? [...new Set([...base, k.id])] : base.filter((x) => x !== k.id));
                        }} />
                      <span className="nm">{k.nome} <Campanella k={k} size={12} /></span>
                      <span className="st">{valutazionePronta(k) ? "da rivedere" : inValutazione(k) ? "in valutazione" : s.faseL}</span>
                      <span className="alt riga" style={{ gap: 7, color: s.colore }}><Pallino v={s.segnale} /> {s.etichetta}</span>
                    </label>
                  );
                })}
              </div>
            </>
          )}
        </section>

        {esiti.map(({ k, s }) => (
          <section className="card" key={k.id}>
            <header>
              <Pallino v={s.segnale} size={13} />
              <h2>{k.nome}</h2>
              <Campanella k={k} size={15} />
              <span className="sp eyebrow">
                {s.faseL}{s.giorni ? ` · giorno ${int(s.giorni)}` : ""} · valutata {s.dove}
                {k.inizioMeta ? ` · partenza reale ${dataIt(k.inizio)}` : ""}
              </span>
              <span className="tag" style={{ background: `${s.colore}22`, color: s.colore }}>{s.etichettaBreve || s.etichetta}</span>
              <button className="btn ghost mini" onClick={() => chiedi({
                titolo: "Registrare l'intervento?",
                testo: `Segno che oggi hai modificato “${k.nome}”. Da adesso il semaforo la lascia in attesa per ${ATTESA_MODIFICA} giorni: dopo un intervento l'algoritmo ritara la consegna, e giudicare sui dati di quei giorni significa giudicare la transizione.`,
                etichetta: "Registra",
                campo: { label: "Cosa hai cambiato", valore: "" },
                azione: (nota) => {
                  const idNota = uid();
                  upd("campagne", (l) => l.map((x) => (x.id === k.id
                    ? { ...x, ultimaModifica: oggi(), notaModifica: nota || "", notaModificaId: idNota, modificaVista: "no" } : x)));
                  /* La stessa modifica finisce nel Diario: lì resta la storia degli interventi,
                     qui solo l'ultimo. Il tipo lo deduco da come l'hai scritta. */
                  const t = String(nota || "").trim().toLowerCase();
                  const tipo = /budget|spesa|giornalier/.test(t) ? "budget"
                    : /annunc|creativ|video|copy|immagin|ad\d/.test(t) ? "creativita"
                      : /pubblic|target|interess|lookalike|audience/.test(t) ? "pubblico"
                        : /offert|prezzo|promo|sconto/.test(t) ? "offerta"
                          : /pixel|tracc|conversion api|capi/.test(t) ? "tracciamento" : "altro";
                  upd("annotazioni", (l) => [{
                    id: idNota, clienteId: k.clienteId, campagnaId: k.id, data: oggi(), tipo,
                    testo: `${nota} — su “${k.nome}”. Giudizio sospeso fino al ${dataItA(sommaGiorni(oggi(), ATTESA_MODIFICA))}.`,
                  }, ...l]);
                },
              })}>{k.ultimaModifica ? "Nuova modifica" : "Ho fatto una modifica"}</button>

              {valutazionePronta(k) && (
                <button className="btn" onClick={() => upd("campagne", (l) => l.map((x) => (x.id === k.id
                  ? { ...x, modificaVista: "si" } : x)))}>Ho rivisto</button>
              )}

              {k.ultimaModifica && (
                <>
                  <button className="btn ghost mini" title="Correggi testo o data dell'ultima modifica"
                    onClick={() => chiedi({
                      titolo: "Correggere la modifica registrata?",
                      testo: `Registrata il ${dataItA(k.ultimaModifica)}. Cambiando il testo aggiorno anche la nota nel Diario, se c'è ancora.`,
                      etichetta: "Salva",
                      campo: { label: "Cosa hai cambiato", valore: k.notaModifica || "" },
                      azione: (nota) => {
                        upd("campagne", (l) => l.map((x) => (x.id === k.id ? { ...x, notaModifica: nota } : x)));
                        if (k.notaModificaId) upd("annotazioni", (l) => l.map((a) => (a.id === k.notaModificaId
                          ? { ...a, testo: `${nota} — su “${k.nome}”. Giudizio sospeso fino al ${dataItA(sommaGiorni(k.ultimaModifica, ATTESA_MODIFICA))}.` } : a)));
                      },
                    })}>Correggi</button>

                  <button className="btn ghost mini" title="Togli la modifica e riprendi a giudicare"
                    onClick={() => chiedi({
                      titolo: "Eliminare la modifica registrata?",
                      testo: `Tolgo la modifica del ${dataItA(k.ultimaModifica)}${k.notaModifica ? ` — ${k.notaModifica}` : ""}: il semaforo torna a giudicare subito questa campagna.\n\nScegli se tenere la nota nel Diario, che è la storia dell'intervento, o cancellare anche quella.`,
                      etichetta: "Togli, tieni il Diario",
                      azione: () => upd("campagne", (l) => l.map((x) => (x.id === k.id
                        ? { ...x, ultimaModifica: "", notaModifica: "", notaModificaId: "" } : x))),
                      etichetta2: "Togli anche dal Diario",
                      azione2: () => {
                        upd("campagne", (l) => l.map((x) => (x.id === k.id
                          ? { ...x, ultimaModifica: "", notaModifica: "", notaModificaId: "" } : x)));
                        if (k.notaModificaId) upd("annotazioni", (l) => l.filter((a) => a.id !== k.notaModificaId));
                      },
                    })}>Elimina</button>
                </>
              )}
              <button className="btn ghost mini" onClick={() => apri("campagne", data.campagne.find((x) => x.id === k.id))}>Apri</button>
            </header>

            <div className="kpis" style={{ padding: 16 }}>
              <div className="kpi"><IcoKpi t="CPL" /><div className="eyebrow">Costo per risultato</div>
                <div className="v">{s.costo ? eur(s.costo) : "—"}</div>
                <div className="s">{s.kpi ? `obiettivo ${eur(s.kpi)}` : "KPI non impostato"}</div></div>
              <div className="kpi"><IcoKpi t="Percentuale" /><div className="eyebrow">Rispetto al KPI</div>
                <div className="v" style={{ color: s.rapporto > 150 ? segna("rosso").c : s.rapporto > 100 ? segna("giallo").c : s.rapporto ? segna("verde").c : undefined }}>
                  {s.rapporto ? `${fmt(s.rapporto, 0)}%` : "—"}
                </div>
                <div className="s">rosso oltre 150 · giallo da 101</div></div>
              <div className="kpi"><IcoKpi t="Erogazione" /><div className="eyebrow">Budget consumato</div>
                <div className="v" style={{ color: s.pacing === null ? undefined : s.pacing ? segna("verde").c : segna("giallo").c }}>
                  {s.pacing === null ? "—" : s.pacing ? "pieno" : "parziale"}
                </div>
                <div className="s">{s.pacing ? "scaling verticale possibile" : s.pacing === false ? "solo orizzontale" : "manca budget o date"}</div></div>
              <div className="kpi"><IcoKpi t="Spesa" /><div className="eyebrow">Speso</div>
                <div className="v">{eur0(k.d.sp)}</div>
                <div className="s">{k.d.ld ? `${int(k.d.ld)} lead` : `${int(k.d.cv)} conversioni`}{s.checkpoint ? ` · ${s.checkpoint}` : ""}</div></div>
            </div>

            {(() => {
              const storia = (data.annotazioni || []).filter((a) => a.campagnaId === k.id)
                .sort((a, b) => (a.data < b.data ? 1 : -1)).slice(0, 4);
              if (!storia.length) return null;
              return (
                <div className="nota" style={{ display: "grid", gap: 7 }}>
                  <span className="eyebrow">Interventi registrati</span>
                  {storia.map((a) => (
                    <button key={a.id} className="riga" style={{ gap: 10, textAlign: "left", background: "transparent", padding: 0 }}
                      onClick={() => apri("annotazioni", a)}>
                      <span className="mono sub" style={{ flex: "none" }}>{dataIt(a.data)}</span>
                      <span className="tag">{et(TIPI_NOTA, a.tipo)}</span>
                      <span style={{ flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.testo}</span>
                    </button>
                  ))}
                </div>
              );
            })()}

            <article className="al" data-g={s.segnale === "rosso" ? "critica" : s.segnale === "giallo" ? "attenzione" : "nota"}>
              <div className="meta">
                <span className="tag" style={{ background: `${s.colore}22`, color: s.colore }}>{s.etichetta}</span>
                <span className="misura">{s.faseL}</span>
              </div>
              <h3>{s.perche}</h3>
              <dl><dt>Cosa fare</dt><dd>{s.azione}</dd></dl>
            </article>

            {/* Analisi 30/7/3: la stessa campagna letta su tre orizzonti.
                Il totale dice com'è andata, le finestre corte dicono come sta andando. */}
            {(() => {
              const ws = ORIZZONTI.map((o) => {
                const w = finestraMobile(data, k, o.n);
                return { o, w, f: semaforoFinestra(k, s.kpi, w) };
              }).filter((x) => x.w);
              if (!ws.length) return null;
              const conCosto = ws.filter((x) => x.f.costo > 0);
              const c30 = (ws.find((x) => x.o.n === 30) || {}).f;
              const c7 = (ws.find((x) => x.o.n === 7) || {}).f;
              const c3 = (ws.find((x) => x.o.n === 3) || {}).f;
              const lettura = (() => {
                /* Su una campagna giovane le finestre si sovrappongono: dirlo è più
                   utile che inventare una tendenza fra numeri identici. */
                const distinte = new Set(ws.map((x) => `${x.w.dal}|${x.w.al}`));
                const lunga = ws[0].w;
                const premessa = distinte.size < ORIZZONTI.length
                  ? `La campagna è partita il ${dataItA(k.inizio)} e ha ${int(lunga.giorni)} ${lunga.giorni === 1 ? "giorno" : "giorni"} di vita, quindi le finestre più lunghe coincidono: si separeranno da sole man mano che cresce. `
                  : "";
                if (distinte.size === 1) return `${premessa}Non c'è ancora una direzione da leggere.`;
                if (conCosto.length < 2) return `${premessa}Servono almeno due finestre con risultati per leggere una direzione.`;
                const costi = conCosto.map((x) => Math.round(x.f.costo * 100));
                if (new Set(costi).size === 1) {
                  return `${premessa}Il costo per risultato è identico su tutte le finestre con dati, ${eur(conCosto[0].f.costo)}: i giorni recenti non si discostano dalla media del periodo.`;
                }
                if (c30 && c7 && c3 && c30.costo && c7.costo && c3.costo) {
                  if (c3.costo > c7.costo && c7.costo > c30.costo) {
                    return `${premessa}Il costo sale avvicinandosi a oggi: ${eur(c30.costo)} sui trenta giorni, ${eur(c7.costo)} sugli ultimi sette, ${eur(c3.costo)} sugli ultimi tre. Il totale è ancora tenuto su dai giorni buoni di tre settimane fa: è il declino lento, quello che non fa rumore finché non è tardi.`;
                  }
                  if (c3.costo < c7.costo && c7.costo < c30.costo) {
                    return `${premessa}Il costo scende avvicinandosi a oggi: ${eur(c30.costo)} sui trenta giorni, ${eur(c7.costo)} sugli ultimi sette, ${eur(c3.costo)} sugli ultimi tre. Quello che hai cambiato sta funzionando, e il totale lo mostrerà con qualche giorno di ritardo.`;
                  }
                }
                const peggio = conCosto.slice().sort((a, b) => b.f.rapporto - a.f.rapporto)[0];
                return `${premessa}Nessuna direzione netta fra le tre finestre: il costo oscilla. La finestra più lontana dall'obiettivo è ${peggio.o.l.toLowerCase()}, al ${fmt(peggio.f.rapporto, 0)}% del KPI.`;
              })();
              const nomeCosto = s.kpiNome === "CPL" ? "CPL" : "CPA";
              return (
                <>
                  <div className="wrap"><table className="tbl">
                    <thead><tr>
                      <th>Finestra</th><th>Periodo</th><th className="r">Speso</th>
                      <th className="r">{s.kpiNome === "CPL" ? "Lead" : "Conv."}</th>
                      <th className="r">{nomeCosto}</th><th className="r">% KPI</th>
                      <th className="r">Erogazione</th><th>Segnale</th>
                    </tr></thead>
                    <tbody>
                      {ws.map(({ o, w, f }) => (
                        <tr key={o.n}>
                          <td>
                            <b>{o.l}</b>
                            {w.tagliata && <div className="sub">dalla partenza: {int(w.giorni)} {w.giorni === 1 ? "giorno" : "giorni"}</div>}
                            {!w.tagliata && w.parziali > 0 && <div className="sub">{w.parziali === 1 ? "una finestra parziale" : `${w.parziali} finestre parziali`}</div>}
                          </td>
                          <td className="sub mono">{dataIt(w.dal)} – {dataIt(w.al)}</td>
                          <td className="r">{eur(w.sp)}</td>
                          <td className="r">{f.conv ? int(f.conv) : "—"}</td>
                          <td className="r">{f.costo ? eur(f.costo) : "—"}</td>
                          <td className={"r " + (f.rapporto > 150 ? "neg" : f.rapporto && f.rapporto <= 100 ? "pos" : "")}>
                            {f.rapporto ? `${fmt(f.rapporto, 0)}%` : "—"}
                          </td>
                          <td className="r">{f.quota === null ? "—" : `${fmt(f.quota * 100, 0)}%`}</td>
                          <td>
                            <span className="riga" style={{ gap: 7, color: f.colore }}>
                              <Pallino v={f.segnale} /> {f.etichetta}
                            </span>
                            <div className="sub">{f.perche}</div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table></div>
                  <div className="nota">
                    <b>Come leggerle.</b> {lettura}
                    {" "}Le tre finestre finiscono ieri ({dataItA(sommaGiorni(oggi(), -1))}): il giorno in corso è parziale e
                    affiancarlo a giorni interi sposterebbe i costi senza motivo. I numeri sono ricostruiti dalle finestre
                    compilate, in proporzione ai giorni che cadono dentro l'intervallo.
                  </div>
                </>
              );
            })()}

            {(() => {
              const kpi = s.kpi;
              const figli = (attivo ? S.creativi : data.creativi).filter((c) => c.campagnaId === k.id);
              if (!figli.length) return (
                <div className="nota">
                  Nessun ad set o annuncio registrato per questa campagna: il giudizio resta al livello della campagna.
                  Importandoli, qui sotto compare anche quale spegnere e quale scalare.
                </div>
              );
              const tutti = figli.map((c) => ({ c, f: semaforoFiglio(c, kpi, k.d.sp) }))
                .sort((a, b) => b.f.sp - a.f.sp);
              /* Un ad set o annuncio senza spesa né impression è un residuo di importazioni
                 diverse: mostrarlo accanto a quello che lavora genera solo confusione. */
              const vuoti = tutti.filter((x) => !x.f.sp && !(x.c.impression || 0));
              const valutati = tutti.filter((x) => x.f.sp > 0 || (x.c.impression || 0) > 0);
              const adset = valutati.filter((x) => x.c.tipo === "adset");
              const annunci = valutati.filter((x) => x.c.tipo !== "adset");
              const rossi = valutati.filter((x) => x.f.segnale === "rosso");
              const verdi = valutati.filter((x) => x.f.segnale === "verde");
              const grigi = valutati.filter((x) => x.f.segnale === "grigio");
              const nomeAdset = (id) => (figli.find((x) => x.id === id) || {}).nome || "";

              /* Che cosa farne, detto in una frase sola */
              /* Il costo della campagna senza il peso morto: è il numero che dice
                 se spegnere un pezzo basta, o se va spenta tutta. */
              const senzaRossi = (() => {
                if (!rossi.length) return null;
                const sp = k.d.sp - rossi.reduce((a, x) => a + x.f.sp, 0);
                const cv = (k.d.ld || k.d.cv) - rossi.reduce((a, x) => a + x.f.conv, 0);
                if (sp <= 0 || cv <= 0) return null;
                const costo = sp / cv;
                return { costo, perc: s.kpi ? (costo / s.kpi) * 100 : 0, sp, cv,
                  vivi: valutati.filter((x) => !["rosso", "grigio"].includes(x.f.segnale)) };
              })();

              const consiglio = () => {
                if (s.segnale === "rosso" && senzaRossi && senzaRossi.vivi.length) {
                  const r = senzaRossi;
                  const resta = r.perc > 150 ? "resta rossa" : r.perc > 100 ? "diventa gialla" : "rientra nel KPI";
                  const dopo = r.perc > 150
                    ? "Non basta: anche senza quelli il costo resta oltre il 150% del KPI, quindi la campagna va spenta tutta."
                    : r.perc > 100
                      ? "Non è un via libera a scalare: a quel punto è gialla, cioè si lascia correre e si guarda al prossimo controllo. Ma è meglio che spegnere qualcosa che funziona."
                      : "A quel punto è dentro l'obiettivo, e la campagna si può tenere.";
                  return `Prima di spegnere tutto, togli il peso morto: ${rossi.map((x) => `“${x.c.nome}”`).join(", ")} ${rossi.length === 1 ? "ha" : "hanno"} consumato ${eur0(k.d.sp - r.sp)} senza portare risultati. ` +
                    `Restano ${r.vivi.map((x) => `“${x.c.nome}”`).join(", ")}: contando solo ${r.vivi.length === 1 ? "quello" : "quelli"}, il costo per risultato è ${eur(r.costo)}, cioè ${fmt(r.perc, 0)}% del KPI, e la campagna ${resta}. ${dopo} ` +
                    `Spegnendo il resto, il budget si concentra su chi converte invece di essere diviso con chi non lo fa. Registra l\u2019intervento con il pulsante qui sopra: il giudizio si sospende per ${ATTESA_MODIFICA} giorni, il tempo che l\u2019algoritmo ritari la consegna.`;
                }
                if (s.segnale === "rosso") {
                  if (rossi.length && verdi.length)
                    return `Non spegnere tutta la campagna: ${rossi.length === 1 ? "è" : "sono"} ${rossi.map((x) => `“${x.c.nome}”`).join(", ")} a trascinarla giù, mentre ${verdi.map((x) => `“${x.c.nome}”`).join(", ")} ${verdi.length === 1 ? "sta" : "stanno"} dentro il KPI. Spegni ${rossi.length === 1 ? "quello" : "quelli"} e lascia correre il resto.`;
                  if (rossi.length)
                    return `Sono rossi ${rossi.map((x) => `“${x.c.nome}”`).join(", ")}, che insieme fanno ${eur0(rossi.reduce((a, x) => a + x.f.sp, 0))} dei ${eur0(k.d.sp)} spesi: qui non c'è un pezzo da salvare, la campagna va spenta.`;
                  return "Nessun ad set o annuncio ha abbastanza spesa per essere giudicato da solo: la decisione resta sulla campagna.";
                }
                if (s.segnale === "verde") {
                  const migliore = verdi[0] || valutati.find((x) => x.f.segnale !== "grigio");
                  if (!migliore) return "Il verde è della campagna nel suo insieme: nessun singolo annuncio ha ancora spesa sufficiente per sapere chi lo sta producendo.";
                  const peggiori = valutati.filter((x) => ["rosso", "giallo"].includes(x.f.segnale));
                  return `Il risultato lo porta “${migliore.c.nome}”: ${eur0(migliore.f.sp)} spesi, ${int(migliore.f.conv)} risultati, ${fmt(migliore.f.rapporto, 0)}% del KPI. È lì che va il budget in più.` +
                    (peggiori.length ? ` Prima però togli ${peggiori.map((x) => `“${x.c.nome}”`).join(", ")}: scalare una campagna che contiene un annuncio sopra il KPI significa pagare di più anche quello.` : "");
                }
                if (s.segnale === "giallo") {
                  if (rossi.length)
                    return `La campagna è gialla ma contiene ${rossi.length === 1 ? "un elemento rosso" : `${rossi.length} elementi rossi`}: ${rossi.map((x) => `“${x.c.nome}”`).join(", ")}. Spegnere ${rossi.length === 1 ? "quello" : "quelli"} è l'unico intervento ammesso sul giallo, e spesso riporta la campagna nel KPI senza toccare altro.`;
                  return "Nessun elemento rosso dentro: il giallo è distribuito, quindi non c'è niente da spegnere. Si aspetta.";
                }
                return "La campagna non è ancora giudicabile, e nemmeno i suoi elementi.";
              };

              /* Riga è un componente: React la disegna dopo, quando la valuta è già tornata quella di prima */
              const vRiga = VALUTA;
              const Riga = ({ x, liv }) => (
                <tr>
                  <td style={{ paddingLeft: 14 + liv * 20 }}>
                    <div className="riga" style={{ gap: 8 }}>
                      {liv > 0 && <span className="sub">↳</span>}
                      <span style={{ width: 9, height: 9, borderRadius: "50%", flex: "none",
                        background: (SEGNALI.find((g) => g.v === x.f.segnale) || {}).c }} />
                      <div>
                        <b>{x.c.nome}</b>
                        <div className="sub">{x.f.perche}</div>
                      </div>
                    </div>
                  </td>
                  <td className="sub">{et(TIPI_CREATIVO, x.c.tipo)}</td>
                  <td className="r">{eurDi0(x.f.sp, vRiga)}</td>
                  <td className="r sub">{x.f.quota ? pctv(x.f.quota * 100, 0) : "—"}</td>
                  <td className="r">{x.f.conv ? int(x.f.conv) : "—"}</td>
                  <td className="r">{x.f.costo ? eurDi(x.f.costo, vRiga) : "—"}</td>
                  <td className={"r " + (x.f.rapporto > 150 ? "neg" : x.f.rapporto && x.f.rapporto <= 100 ? "pos" : "")}>
                    {x.f.rapporto ? `${fmt(x.f.rapporto, 0)}%` : "—"}
                  </td>
                  <td className="r">{x.f.ctr ? pctv(x.f.ctr) : "—"}</td>
                </tr>
              );

              return (
                <>
                  <div className="wrap">
                    <table className="tbl">
                      <thead><tr>
                        <th>Ad set e annunci</th><th>Livello</th><th className="r">Speso</th><th className="r">Quota</th>
                        <th className="r">Risultati</th><th className="r">Costo</th><th className="r">Sul KPI</th><th className="r">CTR link</th>
                      </tr></thead>
                      <tbody>
                        {adset.map((a) => (
                          <React.Fragment key={a.c.id}>
                            <Riga x={a} liv={0} />
                            {annunci.filter((n) => n.c.adsetId === a.c.id).map((n) => <Riga x={n} liv={1} key={n.c.id} />)}
                          </React.Fragment>
                        ))}
                        {annunci.filter((n) => !n.c.adsetId || !adset.some((a) => a.c.id === n.c.adsetId))
                          .map((n) => <Riga x={n} liv={0} key={n.c.id} />)}
                      </tbody>
                    </table>
                  </div>
                  {senzaRossi && senzaRossi.vivi.length > 0 && s.segnale !== "verde" && (
                    <div className="kpis" style={{ padding: "0 16px 14px" }}>
                      <div className="kpi">
                        <IcoKpi t="CPL" /><div className="eyebrow">Costo ora</div>
                        <div className="v" style={{ color: (SEGNALI.find((g) => g.v === s.segnale) || {}).c }}>{s.costo ? eur(s.costo) : "—"}</div>
                        <div className="s">{s.rapporto ? `${fmt(s.rapporto, 0)}% del KPI` : "—"}</div>
                      </div>
                      <div className="kpi">
                        <IcoKpi t="CPL" /><div className="eyebrow">Costo senza i rossi</div>
                        <div className="v" style={{ color: senzaRossi.perc > 150 ? "#D8455F" : senzaRossi.perc > 100 ? "#C99A3A" : "#3FA45C" }}>
                          {eur(senzaRossi.costo)}
                        </div>
                        <div className="s">{fmt(senzaRossi.perc, 0)}% del KPI · {int(senzaRossi.cv)} risultati su {eur0(senzaRossi.sp)}</div>
                      </div>
                      <div className="kpi">
                        <IcoKpi t="Spesa" /><div className="eyebrow">Speso dal peso morto</div>
                        <div className="v">{eur0(k.d.sp - senzaRossi.sp)}</div>
                        <div className="s">{pctv(((k.d.sp - senzaRossi.sp) / k.d.sp) * 100, 0)} del budget, zero risultati</div>
                      </div>
                    </div>
                  )}

                  {vuoti.length > 0 && (
                    <div className="nota" style={{ paddingTop: 10 }}>
                      {vuoti.length === 1 ? "Un elemento senza dati non è mostrato" : `${vuoti.length} elementi senza dati non sono mostrati`}
                      {" "}({vuoti.map((x) => `“${x.c.nome}”`).join(", ")}): nessuna spesa e nessuna impression, di solito
                      sono residui di importazioni fatte con nomi diversi. Li trovi nella tabella Ad set e annunci, dove puoi eliminarli.
                    </div>
                  )}

                  <article className="al" data-g="nota">
                    <div className="meta">
                      <span className="tag">dentro la campagna</span>
                      <span className="misura">
                        {verdi.length} verdi · {valutati.filter((x) => x.f.segnale === "giallo").length} gialli · {rossi.length} rossi
                        {grigi.length ? ` · ${grigi.length} senza spesa sufficiente` : ""}
                      </span>
                    </div>
                    <h3>Su cosa intervenire</h3>
                    <dl><dt>Dettaglio</dt><dd>{consiglio()}</dd></dl>
                  </article>
                </>
              );
            })()}
          </section>
        ))}

        <section className="card">
          <header><h2>Le regole, per intero</h2><span className="sp eyebrow">due indicatori, tre risposte</span></header>
          <div className="wrap">
            <table className="tbl">
              <thead><tr><th>Condizione</th><th>Segnale</th><th>Azione</th></tr></thead>
              <tbody>
                <tr><td>Nessuna conversione, oppure costo oltre il 150% del KPI</td>
                  <td><span className="riga" style={{ gap: 7 }}><Pallino v="rosso" /> Rosso</span></td><td>Spegni</td></tr>
                <tr><td>Costo fra il 101% e il 150% del KPI</td>
                  <td><span className="riga" style={{ gap: 7 }}><Pallino v="giallo" /> Giallo</span></td><td>Non fare nulla</td></tr>
                <tr><td>Costo nel KPI ma budget non consumato del tutto</td>
                  <td><span className="riga" style={{ gap: 7 }}><Pallino v="giallo" /> Giallo</span></td><td>Nessun aumento di budget</td></tr>
                <tr><td>Costo pari o sotto il KPI, con budget consumato</td>
                  <td><span className="riga" style={{ gap: 7 }}><Pallino v="verde" /> Verde</span></td><td>Scala</td></tr>
                <tr><td>Checkpoint di fase non ancora superato</td>
                  <td><span className="riga" style={{ gap: 7 }}><Pallino v="grigio" /> Grigio</span></td><td>Aspetta</td></tr>
              </tbody>
            </table>
          </div>
          <div className="nota">
            I checkpoint seguono le fasi che già usi. <b>Proof of concept</b> è la fase Test: serve una spesa pari almeno al KPI
            prima di esprimersi, solida al doppio. <b>Validation</b> è la fase Matura: sette giorni di attesa prima del giudizio.
            <b> Scaling e mantenimento</b> è la fase Attiva, dove la campagna si guarda ogni giorno e l'aumento di budget richiede
            il budget pieno speso negli ultimi tre giorni consecutivi.
          </div>
        </section>
      </>
    ));
  };

  /* ---------- Pixel facile ---------- */
  const SezionePixel = () => {
    const ev = EVENTI_PIXEL.find((x) => x.v === pxEvento) || EVENTI_PIXEL[0];
    const val = (k) => (pxVal[k] || "").trim();
    const usati = ev.p.filter((k) => val(k));

    const codice = (() => {
      const righe = usati.map((k) => {
        const v = val(k);
        const numerico = ["value", "num_items", "predicted_ltv"].includes(k) && /^-?[\d.,\s]+$/.test(v);
        const lista = ["content_ids", "contents"].includes(k) && /^[[{]/.test(v);
        /* 1.500,00 → 1500.00: il punto italiano separa le migliaia, la virgola i decimali */
        const pulito = v.replace(/\s/g, "").replace(/\.(?=\d{3}\b)/g, "").replace(",", ".");
        const scritto = numerico ? pulito : lista ? v : `'${v.replace(/'/g, "\\'")}'`;
        return `  ${k}: ${scritto}`;
      });
      const corpo = righe.length ? `, {\n${righe.join(",\n")}\n}` : "";
      return `<script>\n  fbq('track', '${ev.v}'${corpo});\n</script>`;
    })();

    const copia = () => {
      const t = document.getElementById("px-out");
      if (t) {
        const sel = window.getSelection();
        const r = document.createRange();
        r.selectNodeContents(t); sel.removeAllRanges(); sel.addRange(r);
        try { document.execCommand("copy"); } catch (e) { /* resta selezionato */ }
        sel.removeAllRanges();
      }
      if (navigator.clipboard) navigator.clipboard.writeText(codice).catch(() => {});
      setPxCopiato(true);
      setTimeout(() => setPxCopiato(false), 2200);
    };

    return (
      <>
        <section className="card">
          <header>
            <h2>1. L'evento da tracciare</h2>
            <span className="sp eyebrow">{EVENTI_PIXEL.length} eventi standard di Meta</span>
            {usati.length > 0 && <button className="btn ghost mini" onClick={() => setPxVal({})}>Svuota i parametri</button>}
          </header>
          <div className="pannello-filtri">
            <div className="f">
              <label htmlFor="px-ev">Evento</label>
              <select id="px-ev" value={pxEvento} onChange={(e) => setPxEvento(e.target.value)}>
                {EVENTI_PIXEL.map((x) => <option key={x.v} value={x.v}>{x.l} · {x.v}</option>)}
              </select>
            </div>
            {attivo && (
              <div className="f">
                <label>Scorciatoie da {attivo.nome}</label>
                <div className="riga" style={{ gap: 7, paddingTop: 6, flexWrap: "wrap" }}>
                  {ev.p.includes("currency") && (
                    <button className="btn ghost mini" onClick={() => setPxVal({ ...pxVal, currency: "EUR" })}>Valuta EUR</button>
                  )}
                  {ev.p.includes("value") && attivo.aov > 0 && (
                    <button className="btn ghost mini" onClick={() => setPxVal({ ...pxVal, value: String(attivo.aov) })}>
                      Valore {eur(attivo.aov)}
                    </button>
                  )}
                  {ev.p.includes("content_category") && attivo.settore && (
                    <button className="btn ghost mini" onClick={() => setPxVal({ ...pxVal, content_category: attivo.settore })}>
                      Area “{attivo.settore}”
                    </button>
                  )}
                </div>
              </div>
            )}
          </div>
          <div className="nota">{ev.d}</div>
        </section>

        <section className="card">
          <header>
            <h2>2. I parametri</h2>
            <span className="sp eyebrow">
              {ev.p.length ? `${usati.length} di ${ev.p.length} compilati · tutti facoltativi` : "questo evento non ne prevede"}
            </span>
          </header>
          {ev.p.length === 0 ? (
            <div className="nota" style={{ paddingTop: 14 }}>
              {ev.l} si traccia da solo: non accetta parametri. Copia il codice qui sotto e mettilo dove l'evento avviene.
            </div>
          ) : (
            <div className="pannello-filtri">
              {ev.p.map((k) => {
                const P = PARAM_PIXEL[k];
                return (
                  <div className="f" key={k}>
                    <label htmlFor={`px-${k}`}>{P.l} <span className="sub">· {k}</span></label>
                    {P.opts ? (
                      <select id={`px-${k}`} value={val(k)} onChange={(e) => setPxVal({ ...pxVal, [k]: e.target.value })}>
                        <option value="">— nessuna —</option>
                        {VALUTE_PIXEL.map((c) => <option key={c} value={c}>{c}</option>)}
                      </select>
                    ) : (
                      <input id={`px-${k}`} value={val(k)} placeholder={P.ph}
                        onChange={(e) => setPxVal({ ...pxVal, [k]: e.target.value })} />
                    )}
                    <span className="hint">{P.d}</span>
                  </div>
                );
              })}
            </div>
          )}
        </section>

        <section className="card">
          <header>
            <h2>3. Il codice</h2>
            <span className="sp eyebrow">{ev.l} · {usati.length ? `${usati.length} parametri` : "senza parametri"}</span>
            <button className="btn" onClick={copia}>{pxCopiato ? "Copiato" : "Copia"}</button>
          </header>
          <pre className="px-out"><code id="px-out">{codice}</code></pre>
          <div className="nota">
            Incolla questo blocco nel punto del sito in cui l'evento avviene: nella pagina di ringraziamento per
            l'acquisizione contatto, in quella di conferma per l'acquisto. Deve stare dopo il codice base del pixel,
            altrimenti <b>fbq</b> non esiste ancora e l'evento non parte.
          </div>
        </section>

        <section className="card">
          <header><h2>Dove finisce questo numero</h2></header>
          <div className="nota" style={{ paddingTop: 14 }}>
            L'evento che tracci qui è lo stesso che poi leggi nelle colonne Lead e Conversioni dell'Advertising, e su cui
            si regge il CPL. Un evento piazzato nel posto sbagliato — per esempio su tutte le pagine invece che sulla
            conferma — gonfia i numeri senza che nulla lo segnali: i dati restano coerenti con sé stessi e sbagliati.
            Dopo averlo installato, fai una conversione di prova e controlla che ne arrivi una sola.
          </div>
        </section>
      </>
    );
  };

  /* ---------- calendario ---------- */
  const SezioneCalendario = () => {
    const perno = calData ? new Date(calData + "T00:00:00") : new Date(oggi() + "T00:00:00");
    const iso = isoLocale;
    const sposta = (n) => {
      const d = new Date(perno);
      if (calVista === "giorno") d.setDate(d.getDate() + n);
      else if (calVista === "settimana") d.setDate(d.getDate() + n * 7);
      else if (calVista === "mese") d.setMonth(d.getMonth() + n);
      else d.setFullYear(d.getFullYear() + n);
      setCalData(iso(d));
    };
    const lunedi = (d) => {
      const x = new Date(d);
      x.setDate(x.getDate() - ((x.getDay() + 6) % 7));
      return x;
    };
    const colore = (id) => (data.categorie.find((c) => c.id === id) || {}).colore || "";

    /* Tutto quello che ha una data diventa una voce del calendario */
    const voci = (() => {
      const out = [];
      const dentro = (x) => !attivo || ambito.includes(x.clienteId) || !x.clienteId;
      if (calFonti.eventi) data.eventi.filter(dentro).forEach((e) => {
        let d = e.data;
        const fine = e.dataFine && e.dataFine > e.data ? e.dataFine : e.data;
        let g = 0;
        while (d <= fine && g < 120) {
          out.push({ data: d, titolo: e.titolo, tipo: "eventi", colore: colore(e.categoriaId) || "#8B5CF6",
            clienteId: e.clienteId, rec: e, piu: fine !== e.data,
            ora: d === e.data ? (e.ora || "") : "", durata: e.durata || 1 });
          d = sommaGiorni(d, 1); g++;
        }
      });
      if (calFonti.task) data.task.filter(dentro).filter((t) => t.scadenza).forEach((t) => {
        const p = PRIORITA.find((x) => x.v === t.priorita) || {};
        out.push({ data: t.scadenza, titolo: t.titolo, tipo: "task", rec: t, clienteId: t.clienteId,
          colore: t.stato === "fatto" ? "#3FA45C" : t.priorita === "urgente" ? "#D8455F" : t.priorita === "alta" ? "#E07A3F" : "#C99A3A",
          nota: `${p.l || ""}${t.stato === "fatto" ? " · fatto" : t.scadenza < oggi() ? " · in ritardo" : ""}` });
      });
      if (calFonti.campagne) m.campagne.filter(dentro).forEach((k) => {
        if (k.inizio) out.push({ data: k.inizio, titolo: `Avvio · ${k.nome}`, tipo: "campagne", colore: "#D8455F", clienteId: k.clienteId });
        if (k.fine && k.stato !== "attiva") out.push({ data: k.fine, titolo: `Chiusa · ${k.nome}`, tipo: "campagne", colore: "#6C6C79", clienteId: k.clienteId });
      });
      if (calFonti.finestre) (data.periodi || []).filter(dentro).forEach((pz) => {
        const k = m.campagne.find((x) => x.id === pz.campagnaId);
        if (!k) return;
        out.push({ data: pz.inizio, titolo: `${et(TIPI_PERIODO, pz.tipo)} · ${k.nome}`, tipo: "finestre", colore: "#4C8DF6", clienteId: pz.clienteId });
      });
      if (calFonti.chiusure) data.chiusure.filter(dentro).forEach((c) => {
        out.push({ data: c.data, titolo: `Chiuso · ${c.nome || c.azienda || "cliente"}`, tipo: "chiusure", colore: "#3FA45C",
          clienteId: c.clienteId, nota: c.valoreContratto ? conValuta(valutaCliente(data, c.clienteId), () => eur0(c.valoreContratto)) : "" });
      });
      return out.filter((x) => x.data);
    })();

    const delGiorno = (d) => voci.filter((v) => v.data === d);
    const nuovoIl = (d, ora = "") => { setCalGiorno(d); apri("eventi", { data: d, ora }, true); };

    const titoloPeriodo = () => {
      if (calVista === "giorno") return dataItA(iso(perno));
      if (calVista === "settimana") {
        const l = lunedi(perno), d = new Date(l); d.setDate(d.getDate() + 6);
        return `${dataIt(iso(l))} – ${dataIt(iso(d))}`;
      }
      if (calVista === "mese") return meseIt(iso(perno).slice(0, 7));
      return perno.getFullYear();
    };

    const Pastiglia = ({ v }) => (
      <button className="cal-voce" style={{ borderLeftColor: v.colore }}
        title={`${v.ora ? `${v.ora} · ` : ""}${v.titolo}${v.nota ? ` — ${v.nota}` : ""}`}
        onClick={() => { if (v.rec) apri(v.tipo === "task" ? "task" : "eventi", v.rec); }}>
        {v.ora && <b className="mono" style={{ marginRight: 5, color: "var(--ink)" }}>{v.ora}</b>}
        <span>{v.titolo}</span>
        {v.nota && <em>{v.nota}</em>}
      </button>
    );

    /* --- mese --- */
    const vistaMese = () => {
      const primo = new Date(perno.getFullYear(), perno.getMonth(), 1);
      const inizio = lunedi(primo);
      const celle = [];
      for (let i = 0; i < 42; i++) {
        const d = new Date(inizio); d.setDate(d.getDate() + i);
        const s = iso(d);
        celle.push({ s, fuori: d.getMonth() !== perno.getMonth(), oggi: s === oggi(), voci: delGiorno(s) });
      }
      return (
        <div className="cal-mese">
          {["lun", "mar", "mer", "gio", "ven", "sab", "dom"].map((g) => <div className="cal-testa" key={g}>{g}</div>)}
          {celle.map((c) => (
            <div className={`cal-cella${c.fuori ? " fuori" : ""}${c.oggi ? " oggi" : ""}`} key={c.s}
              onDoubleClick={() => nuovoIl(c.s)}>
              <div className="n">
                <span>{c.s.slice(8)}</span>
                <button className="piu" onClick={() => nuovoIl(c.s)} aria-label={`Aggiungi il ${dataIt(c.s)}`}>+</button>
              </div>
              {c.voci.slice(0, 4).map((v, i) => <Pastiglia v={v} key={i} />)}
              {c.voci.length > 4 && (
                <button className="cal-altro" onClick={() => { setCalVista("giorno"); setCalData(c.s); }}>
                  altri {c.voci.length - 4}
                </button>
              )}
            </div>
          ))}
        </div>
      );
    };

    /* --- settimana --- */
    const vistaSettimana = () => {
      const l = lunedi(perno);
      const gg = Array.from({ length: 7 }, (_, i) => { const d = new Date(l); d.setDate(d.getDate() + i); return iso(d); });
      return (
        <div className="cal-sett">
          {gg.map((s) => (
            <div className={`cal-colonna${s === oggi() ? " oggi" : ""}`} key={s}>
              <div className="cal-testa">
                {new Date(s + "T00:00:00").toLocaleDateString(LOCALE, { weekday: "short" })} {s.slice(8)}
                <button className="piu" onClick={() => nuovoIl(s)} aria-label={`Aggiungi il ${dataIt(s)}`}>+</button>
              </div>
              <div className="corpo">
                {delGiorno(s).map((v, i) => <Pastiglia v={v} key={i} />)}
                {!delGiorno(s).length && <span className="sub vuoto">—</span>}
              </div>
            </div>
          ))}
        </div>
      );
    };

    /* --- giorno: fascia di testa per ciò che non ha orario, poi le ore --- */
    const vistaGiorno = () => {
      const s = iso(perno);
      const l = delGiorno(s);
      const conOra = l.filter((v) => /^\d{2}:\d{2}$/.test(v.ora || ""));
      const senzaOra = l.filter((v) => !/^\d{2}:\d{2}$/.test(v.ora || ""));
      const ore = conOra.map((v) => parseInt(v.ora.slice(0, 2), 10));
      const da = Math.min(7, ...(ore.length ? ore : [7]));
      const a = Math.max(20, ...(ore.length ? ore.map((h) => h + 1) : [20]));
      const fasce = Array.from({ length: a - da + 1 }, (_, i) => da + i);
      const adesso = oggi() === s ? new Date().getHours() + new Date().getMinutes() / 60 : null;

      return (
        <div className="cal-giorno">
          <div className="tuttoilgiorno">
            <span className="et">Tutto il giorno</span>
            <div className="voci">
              {senzaOra.length
                ? senzaOra.map((v, i) => <Pastiglia v={v} key={i} />)
                : <span className="sub">niente senza orario</span>}
            </div>
          </div>

          <div className="ore">
            {fasce.map((h) => {
              const inFascia = conOra.filter((v) => parseInt(v.ora.slice(0, 2), 10) === h);
              return (
                <div className="fascia" key={h} onDoubleClick={() => nuovoIl(s, `${String(h).padStart(2, "0")}:00`)}>
                  <span className="h">{String(h).padStart(2, "0")}:00</span>
                  <div className="slot">
                    {inFascia.map((v, i) => (
                      <div className="blocco" key={i}
                        style={{ borderLeftColor: v.colore, background: `${v.colore}1A`,
                          minHeight: 26 * Math.max(1, Math.min(4, v.durata || 1)) }}>
                        <button onClick={() => { if (v.rec) apri(v.tipo === "task" ? "task" : "eventi", v.rec); }}>
                          <b>{v.ora}</b> {v.titolo}
                          {v.nota && <em> · {v.nota}</em>}
                        </button>
                      </div>
                    ))}
                    <button className="aggiungi" onClick={() => nuovoIl(s, `${String(h).padStart(2, "0")}:00`)}>+</button>
                  </div>
                  {adesso !== null && Math.floor(adesso) === h && (
                    <span className="adesso" style={{ top: `${(adesso - h) * 100}%` }} />
                  )}
                </div>
              );
            })}
          </div>
        </div>
      );
    };

    /* --- anno: dodici mesi veri, ogni giorno colorato da ciò che contiene --- */
    const vistaAnno = () => {
      const anno = perno.getFullYear();
      return (
        <div className="cal-anno">
          {Array.from({ length: 12 }, (_, mi) => {
            const mm = `${anno}-${String(mi + 1).padStart(2, "0")}`;
            const primo = new Date(anno, mi, 1);
            const ini = lunedi(primo);
            const nelMese = voci.filter((v) => v.data.slice(0, 7) === mm);
            const celle = Array.from({ length: 42 }, (_, i) => {
              const d = new Date(ini); d.setDate(d.getDate() + i);
              const ds = iso(d);
              return { ds, fuori: d.getMonth() !== mi, voci: nelMese.filter((v) => v.data === ds) };
            });
            const ultima = celle.reduce((acc, c, i) => (c.fuori ? acc : i), 0);
            return (
              <div className="mesetto" key={mm}>
                <header>
                  <b onClick={() => { setCalVista("mese"); setCalData(`${mm}-01`); }}>{meseIt(mm).split(" ")[0]}</b>
                  <span className="n mono">{nelMese.length || ""}</span>
                </header>
                <div className="cal-mini">
                  {["L", "M", "M", "G", "V", "S", "D"].map((g, i) => <span className="dow" key={i}>{g}</span>)}
                  {celle.slice(0, Math.ceil((ultima + 1) / 7) * 7).map((c) => (
                    <button className={`gg${c.fuori ? " fuori" : ""}${c.ds === oggi() ? " oggi" : ""}`} key={c.ds}
                      title={c.voci.length ? `${dataIt(c.ds)} · ${c.voci.map((v) => v.titolo).join(" · ")}` : dataIt(c.ds)}
                      onClick={() => { setCalVista("giorno"); setCalData(c.ds); }}>
                      {c.ds.slice(8)}
                      {c.voci.length > 0 && (
                        <span className="punti">
                          {[...new Set(c.voci.map((v) => v.colore))].slice(0, 3).map((col, i) => (
                            <i key={i} style={{ background: col }} />
                          ))}
                        </span>
                      )}
                    </button>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      );
    };

    return (
      <>
        <section className="card">
          <header>
            <h2>{titoloPeriodo()}</h2>
            <span className="sp" />
            <div className="riga" style={{ gap: 6 }}>
              <button className="btn ghost mini" onClick={() => sposta(-1)} aria-label="Indietro">‹</button>
              <button className="btn ghost mini" onClick={() => setCalData(oggi())}>Oggi</button>
              <button className="btn ghost mini" onClick={() => sposta(1)} aria-label="Avanti">›</button>
            </div>
            <select value={calVista} onChange={(e) => setCalVista(e.target.value)} aria-label="Vista del calendario"
              style={{ width: "auto", minWidth: 118 }}>
              {VISTE_CAL.map((v) => <option key={v.v} value={v.v}>{v.l}</option>)}
            </select>
            <button className="btn" onClick={() => nuovoIl(iso(perno))}>+ Evento</button>
          </header>

          <div className="cal-legenda">
            {FONTI_CAL.map((f) => (
              <label key={f.v} className={calFonti[f.v] ? "on" : ""}>
                <input type="checkbox" checked={!!calFonti[f.v]}
                  onChange={(e) => setCalFonti({ ...calFonti, [f.v]: e.target.checked })} />
                <i style={{ background: f.c }} />
                {f.l}
              </label>
            ))}
            <span className="sp" />
            <span className="sub">{voci.length} voci nel periodo caricato</span>
          </div>

          {calVista === "mese" && vistaMese()}
          {calVista === "settimana" && vistaSettimana()}
          {calVista === "giorno" && vistaGiorno()}
          {calVista === "anno" && vistaAnno()}
        </section>

        <section className="card">
          <header>
            <h2>Categorie e colori</h2>
            <span className="sp eyebrow">il significato lo decidi tu</span>
            <button className="btn ghost mini" onClick={() => apri("categorie")}>+ Categoria</button>
          </header>
          <div className="riga" style={{ gap: 8, flexWrap: "wrap", padding: "14px 16px" }}>
            {data.categorie.map((c) => (
              <button className="cal-cat" key={c.id} onClick={() => apri("categorie", c)}>
                <i style={{ background: c.colore }} />
                {c.nome}
              </button>
            ))}
            {!data.categorie.length && <span className="sub">Nessuna categoria: creane una per colorare gli eventi.</span>}
          </div>
          <div className="nota">
            Le categorie valgono per gli eventi che crei tu. Task, campagne, finestre e chiusure hanno colori propri,
            quelli della legenda qui sopra: i task seguono la priorità, dal bordeaux dell'urgente all'ambra della media.
          </div>
        </section>
      </>
    );
  };

  /* ---------- riepilogo mensile ---------- */
  /* Riepilogo mensile riutilizzabile: lo usano sia la panoramica sia la sezione dedicata */
  const righeMensili = (idsClienti) => {
    const dentro = (x) => !idsClienti || idsClienti.includes(x.clienteId);
    const camp = m.campagne.filter(dentro);
    const idCamp = camp.map((k) => k.id);
    const fasi = (data.periodi || []).filter((pz) => idCamp.includes(pz.campagnaId) && pz.inizio);
    const chius = data.chiusure.filter(dentro);
    const vend = data.vendite.filter(dentro);
    const gior = data.giornate.filter(dentro);

    const mesi = [...new Set([
      ...fasi.map((pz) => mese(pz.inizio)),
      ...chius.map((c) => mese(c.data)),
      ...vend.map((v) => v.periodo),
      ...gior.map((g) => mese(g.data)),
    ].filter(Boolean))].sort();

    return mesi.map((mm) => {
      const f = fasi.filter((pz) => mese(pz.inizio) === mm);
      const v = vend.filter((x) => x.periodo === mm);
      const c = chius.filter((x) => mese(x.data) === mm);
      const g = gior.filter((x) => mese(x.data) === mm);
      const spesaFasi = f.reduce((s2, pz) => s2 + (pz.spesa || 0), 0);
      const spesa = spesaFasi > 0 ? spesaFasi : v.reduce((s2, x) => s2 + (x.spesa || 0), 0);
      const sommaG = (k) => g.reduce((s2, x) => s2 + (x[k] || 0), 0);
      const sommaV = (k) => v.reduce((s2, x) => s2 + (x[k] || 0), 0);
      /* Lead e tentativi di contatto si contano cliente per cliente e poi si sommano.
         La fonte è quella che li conosce meglio: le giornate dei venditori, poi il
         ciclo di vendita, infine le finestre pubblicitarie — perché un account che
         vive di lead importati da Meta lasciava la colonna vuota. Dentro lo stesso
         mese due clienti possono usare fonti diverse: prenderne una sola per tutti
         sommerebbe la spesa di entrambi ai lead di uno, e il CPL non direbbe niente. */
      const cliDi = (pz) => pz.clienteId || (camp.find((k) => k.id === pz.campagnaId) || {}).clienteId || "";
      const cls = [...new Set([...f.map(cliDi), ...v.map((x) => x.clienteId), ...g.map((x) => x.clienteId)])];
      let lead = 0, tentativi = 0, nChiamate = 0, nBasi = 0, leadDaFasi = false;
      cls.forEach((cid) => {
        const gc = g.filter((x) => x.clienteId === cid);
        const vc = v.filter((x) => x.clienteId === cid);
        const fc = f.filter((pz) => cliDi(pz) === cid);
        const sG = (kk) => gc.reduce((s2, x) => s2 + (x[kk] || 0), 0);
        const sV = (kk) => vc.reduce((s2, x) => s2 + (x[kk] || 0), 0);
        const lFasi = fc.reduce((s2, pz) => s2 + (pz.lead || 0), 0);
        const l = gc.length ? sG("lead") : (sV("lead") || lFasi);
        if (!gc.length && !sV("lead") && lFasi > 0) leadDaFasi = true;
        const ch = gc.length ? sG("chiamate") : sV("chiamate");
        lead += l;
        tentativi += ch || l;
        if (ch > 0) nChiamate += 1;
        if ((ch || l) > 0) nBasi += 1;
      });
      const demo = g.length ? sommaG("demoSvolte") : sommaV("demoSvolte");
      const chiusure = c.length || (g.length ? sommaG("chiusure") : sommaV("chiusure"));
      const contratto = c.length ? c.reduce((s2, x) => s2 + parseNumIntl(x.valoreContratto), 0)
        : (g.length ? sommaG("valoreContratto") : sommaV("valoreContratto"));
      const incassato = c.length ? c.reduce((s2, x) => s2 + parseNumIntl(x.incassato), 0)
        : (g.length ? sommaG("incassato") : sommaV("incassato"));
      /* Sforzo e resa: quanto incassi per ogni tentativo di contatto. Lo sforzo sono
         le chiamate fatte; dove nessuno le conta valgono i contatti entrati nel funnel,
         e la riga lo dichiara, perché le due basi non sono confrontabili fra loro. */
      return { mm, spesa, daFasi: spesaFasi > 0, lead, leadDaFasi, demo, chiusure, contratto, incassato,
        tentativi, suChiamate: nChiamate > 0 && nChiamate === nBasi,
        resa: tentativi ? incassato / tentativi : 0,
        resaContratto: tentativi ? contratto / tentativi : 0,
        cpl: lead ? spesa / lead : 0, cpChiusura: chiusure ? spesa / chiusure : 0,
        ritorno: spesa ? contratto / spesa : 0, incasso: spesa ? incassato / spesa : 0 };
    });
  };

  const SezioneMesi = () => {
    /* A livello agenzia contano i clienti principali: i sottoclienti si leggono
       entrando nel loro madre. Dentro un cliente vale l'ambito, consolidato o no. */
    const righe = righeMensili(attivo ? ambito : attivi.filter((z) => !z.padreId).map((z) => z.id)).slice().reverse();

    const T = righe.reduce((s, r) => ({
      spesa: s.spesa + r.spesa, lead: s.lead + r.lead, demo: s.demo + r.demo, chiusure: s.chiusure + r.chiusure,
      contratto: s.contratto + r.contratto, incassato: s.incassato + r.incassato,
      tentativi: s.tentativi + (r.tentativi || 0),
    }), { spesa: 0, lead: 0, demo: 0, chiusure: 0, contratto: 0, incassato: 0, tentativi: 0 });
    const Tcpl = T.lead ? T.spesa / T.lead : 0;
    /* Il totale si ricalcola dai totali, non è la media delle righe */
    const Tresa = T.tentativi ? T.incassato / T.tentativi : 0;
    const TcpC = T.chiusure ? T.spesa / T.chiusure : 0;
    const Trit = T.spesa ? T.contratto / T.spesa : 0;
    const Tinc = T.spesa ? T.incassato / T.spesa : 0;

    const esporta = () => {
      const intest = ["Mese", "Spesa ads", "Lead", "Demo svolte", "Chiusure", "CPL", "Costo per chiusura", "Valore contratti", "Incassato", "Tentativi di contatto", "Sforzo/Resa", "Ritorno su contratti", "Ritorno su incassato"];
      const corpo = righe.map((r) => [meseIt(r.mm), r.spesa, r.lead, r.demo, r.chiusure, r.cpl, r.cpChiusura, r.contratto, r.incassato, r.tentativi, r.resa, r.ritorno, r.incasso]);
      const totale = ["Totale", T.spesa, T.lead, T.demo, T.chiusure, Tcpl, TcpC, T.contratto, T.incassato, T.tentativi, Tresa, Trit, Tinc];
      const foglio = XLSX.utils.aoa_to_sheet([intest, ...corpo, [], totale]);
      foglio["!cols"] = intest.map((h) => ({ wch: Math.max(13, h.length + 2) }));
      const cartella = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(cartella, foglio, "Riepilogo mensile");
      const dati = XLSX.write(cartella, { bookType: "xlsx", type: "array" });
      scaricaFile(`riepilogo-mensile-${attivo ? attivo.nome.toLowerCase().replace(/[^a-z0-9]+/g, "-") : "agenzia"}-${oggi()}.xlsx`,
        dati, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    };

    if (!righe.length) return (
      <section className="card">
        <header><h2>Riepilogo mensile</h2></header>
        <div className="empty">
          <b>Ancora nessun mese da riepilogare</b>
          Compare qui appena registri spesa nelle finestre delle campagne, periodi nel ciclo di vendita o schede chiusura.
        </div>
      </section>
    );

    return conValuta(valutaComune(attivo ? ambito : attivi.filter((z) => !z.padreId).map((z) => z.id)), () => (
      <>
        <div className="kpis">
          <div className="kpi"><IcoKpi t="Spesa nel periodo" /><div className="eyebrow">Spesa nel periodo</div><div className="v">{eur0(T.spesa)}</div><div className="s">{righe.length} mesi</div></div>
          <div className="kpi"><IcoKpi t="Chiusure" /><div className="eyebrow">Chiusure</div><div className="v">{int(T.chiusure)}</div><div className="s">{TcpC ? `${eur0(TcpC)} per chiusura` : "—"}</div></div>
          <div className="kpi"><IcoKpi t="Valore contratti" /><div className="eyebrow">Valore contratti</div><div className="v">{eur0(T.contratto)}</div><div className="s">{Trit ? `${fmt(Trit)}× sulla spesa` : "—"}</div></div>
          <div className="kpi"><IcoKpi t="Incassato" /><div className="eyebrow">Incassato</div><div className="v">{eur0(T.incassato)}</div><div className="s">{T.contratto ? `${pctv((T.incassato / T.contratto) * 100, 0)} del firmato` : "—"}</div></div>
        </div>

        <section className="card">
          <header>
            <h2>Mese per mese</h2>
            <span className="sp eyebrow">{attivo ? (haFigli && consolidato ? "gruppo consolidato" : attivo.nome) : "clienti principali"}</span>
            <button className="btn ghost mini" onClick={esporta}>Esporta in Excel</button>
          </header>
          <div className="wrap">
            <table className="tbl">
              <thead><tr>
                <th>Mese</th><th className="r">Spesa ads</th><th className="r">Lead</th><th className="r">CPL</th>
                <th className="r">Demo svolte</th><th className="r">Chiusure</th><th className="r">Costo chiusura</th>
                <th className="r">Valore contratti</th><th className="r">Incassato</th>
                <th className="r">Sforzo/Resa</th><th className="r">Ritorno</th>
              </tr></thead>
              <tbody>
                {righe.map((r) => (
                  <tr key={r.mm}>
                    <td><b>{meseIt(r.mm)}</b>{!r.daFasi && r.spesa > 0 && <div className="sub">spesa dal ciclo di vendita</div>}</td>
                    <td className="r">{r.spesa ? eur0(r.spesa) : "—"}</td>
                    <td className="r">{r.lead ? int(r.lead) : "—"}</td>
                    <td className="r">{r.cpl ? eur(r.cpl) : "—"}</td>
                    <td className="r">{r.demo ? int(r.demo) : "—"}</td>
                    <td className="r">{r.chiusure ? int(r.chiusure) : "—"}</td>
                    <td className="r">{r.cpChiusura ? eur0(r.cpChiusura) : "—"}</td>
                    <td className="r">{r.contratto ? eur0(r.contratto) : "—"}</td>
                    <td className={"r " + (r.contratto && r.incassato < r.contratto ? "neg" : "")}>{r.incassato ? eur0(r.incassato) : "—"}</td>
                    <td className={"r " + (r.resa ? (r.cpl && r.resa >= r.cpl ? "pos" : r.cpl ? "neg" : "") : "")}
                      title={r.resa ? `${eur(r.resa)} incassati per ogni ${r.suChiamate ? "chiamata fatta" : "contatto entrato nel funnel"} (${int(r.tentativi)})${r.cpl ? `, contro ${eur(r.cpl)} di costo per contatto` : ""}` : "Serve il numero di chiamate o di lead del mese"}>
                      {r.resa ? eur(r.resa) : "—"}
                      {r.resa > 0 && <div className="sub">{r.suChiamate ? "per chiamata" : "per contatto"}</div>}
                    </td>
                    <td className={"r " + (r.ritorno ? (r.ritorno >= 1 ? "pos" : "neg") : "")}>{r.ritorno ? `${fmt(r.ritorno)}×` : "—"}</td>
                  </tr>
                ))}
              </tbody>
              <tfoot><tr className="totale">
                <td><b>Totale</b><div className="sub">{righe.length} mesi</div></td>
                <td className="r"><b>{eur0(T.spesa)}</b></td>
                <td className="r"><b>{int(T.lead)}</b></td>
                <td className="r"><b>{Tcpl ? eur(Tcpl) : "—"}</b></td>
                <td className="r"><b>{int(T.demo)}</b></td>
                <td className="r"><b>{int(T.chiusure)}</b></td>
                <td className="r"><b>{TcpC ? eur0(TcpC) : "—"}</b></td>
                <td className="r"><b>{eur0(T.contratto)}</b></td>
                <td className="r"><b>{eur0(T.incassato)}</b></td>
                <td className={"r " + (Tresa ? (Tcpl && Tresa >= Tcpl ? "pos" : Tcpl ? "neg" : "") : "")}>
                  <b>{Tresa ? eur(Tresa) : "—"}</b>
                </td>
                <td className="r"><b>{Trit ? `${fmt(Trit)}×` : "—"}</b></td>
              </tr></tfoot>
            </table>
          </div>
          <div className="nota">
            La spesa arriva dalle finestre delle campagne quando sono compilate, altrimenti da quella dichiarata nel ciclo di vendita:
            nei mesi in cui usa la seconda fonte te lo scrivo sotto il mese. Contratti e incassato vengono dalle schede chiusura,
            e il ritorno è il valore firmato diviso la spesa — non il ROAS della piattaforma.
          </div>
        </section>

        {!attivo && (() => {
          const perCli = attivi.filter((z) => !z.padreId).map((z) => {
            const idc = data.campagne.filter((k) => k.clienteId === z.id && k.archiviata !== "si").map((k) => k.id);
            const sp = (data.periodi || []).filter((pz) => idc.includes(pz.campagnaId)).reduce((s, pz) => s + (pz.spesa || 0), 0)
              || data.vendite.filter((v) => v.clienteId === z.id).reduce((s, v) => s + (v.spesa || 0), 0);
            const cc = data.chiusure.filter((c) => c.clienteId === z.id);
            const rz = righeMensili([z.id]);
            const lead = rz.reduce((s2, r) => s2 + (r.lead || 0), 0);
            return { z, sp, lead, n: cc.length,
              contratto: cc.reduce((s, c) => s + parseNumIntl(c.valoreContratto), 0),
              incassato: cc.reduce((s, c) => s + parseNumIntl(c.incassato), 0) };
          }).filter((x) => x.sp > 0 || x.n > 0).sort((a, b) => b.contratto - a.contratto);
          if (!perCli.length) return null;
          return (
            <section className="card">
              <header><h2>Per cliente</h2><span className="sp eyebrow">clienti principali · tutto il periodo</span></header>
              <div className="wrap">
                <table className="tbl">
                  <thead><tr><th>Cliente</th><th className="r">Spesa</th><th className="r">Lead</th><th className="r">CPL</th><th className="r">Chiusure</th><th className="r">Valore contratti</th><th className="r">Incassato</th><th className="r">Ritorno</th></tr></thead>
                  <tbody>
                    {mapValuta(perCli, (x) => x.z.id, (x) => (
                      <tr key={x.z.id} className="clic" onClick={() => cambiaAccount(x.z.id)}>
                        <td><div className="riga"><Logo src={x.z.logo} nome={x.z.nome} size={26} /><div><b>{x.z.nome}</b>{x.z.padreId && <div className="sub">sottocliente di {nomeCli(x.z.padreId)}</div>}</div></div></td>
                        <td className="r">{x.sp ? eur0(x.sp) : "—"}</td>
                        <td className="r">{x.lead ? int(x.lead) : "—"}</td>
                        <td className="r">{x.lead ? eur(x.sp / x.lead) : "—"}</td>
                        <td className="r">{int(x.n)}</td>
                        <td className="r">{x.contratto ? eur0(x.contratto) : "—"}</td>
                        <td className="r">{x.incassato ? eur0(x.incassato) : "—"}</td>
                        <td className={"r " + (x.sp && x.contratto ? (x.contratto / x.sp >= 1 ? "pos" : "neg") : "")}>
                          {x.sp && x.contratto ? `${fmt(x.contratto / x.sp)}×` : "—"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          );
        })()}
      </>
    ));
  };

  /* ---------- piano di messaggio e Core 4 ---------- */
  const SezioneMessaggio = () => {
    const M = (attivo || {}).msg || {};
    const righe = (t) => String(t || "").split("\n").map((x) => x.trim()).filter(Boolean);
    const ben = righe(M.benefici), obi = righe(M.obiezioni);
    const pieno = M.mercato && M.dolore && M.soluzione && ben.length && M.cosaVendo && M.offerta;
    const sso = M.mercato && M.desiderio
      ? `Aiuto ${M.mercato} a ${M.desiderio}${M.tensione ? `, senza ${M.tensione}` : ""}.`
      : "";

    const core4 = [
      {
        stadio: "1", tit: "Fase 1 · Traffico freddo — problema, soluzione, invito",
        corpo: `${M.dolore || "[il dolore che risolvi, come affermazione o domanda]"} ${M.soluzione || "[la tua soluzione]"}\n\n${(ben.length ? ben : ["[beneficio 1]", "[beneficio 2]", "[beneficio 3]"]).map((b) => `✅ ${b}`).join("\n")}\n\n[invito all'azione]`,
        creativo: "Titolo su blocco di colore pieno",
        titolo: `Solo per ${M.mercato || "[il tuo mercato]"}`,
      },
      {
        stadio: "2", tit: "Fase 2 · Traffico tiepido — informazione utile",
        corpo: `Leggi questo prima di acquistare ${M.cosaVendo || "[cosa vendi]"}.`,
        creativo: "Formato leggero e riconoscibile, tono da contenuto e non da pubblicità",
        titolo: `La guida definitiva all'acquisto di ${M.cosaVendo || "[cosa vendi]"}`,
      },
      {
        stadio: "3", tit: "Fase 3 · Traffico tiepido — chi siamo",
        corpo: `${M.missione || "[la tua missione dichiarata al mercato]"}\n\n${obi[0] ? `So cosa stai pensando: ${obi[0]}. Ecco perché non è un problema qui.` : "[supera l'obiezione più grande]"}\n\n${ben[0] || "[il beneficio più forte]"}`,
        creativo: "Video dimostrativo del prodotto o del servizio",
        titolo: `Perché scegliere ${attivo ? attivo.nome : "[nome azienda]"}?`,
      },
      {
        stadio: "4", tit: "Fase 4 · Traffico caldo — offerta",
        corpo: M.futuro || "[proiezione nel futuro: come si sentirà dopo aver ottenuto il risultato]",
        creativo: "Immagine del prodotto o dell'offerta",
        titolo: `${M.offerta || "[l'offerta]"} — [invito di una parola]`,
      },
    ];

    const creaAnnunci = () => {
      chiedi({
        titolo: "Creare i quattro annunci?",
        testo: "Aggiungo quattro voci in Ad set e annunci, una per fase, con il testo generato nelle note. Non sono collegate a una campagna: la scegli tu aprendole.",
        etichetta: "Crea",
        azione: () => upd("creativi", (l) => [
          ...core4.map((c) => ({
            id: uid(), clienteId: cliAtt, campagnaId: "", tipo: "annuncio", stadio: c.stadio,
            nome: `Core 4 · ${(STADI.find((s) => s.v === c.stadio) || {}).b}`, stato: "pausa",
            spesa: 0, impression: 0, click: 0, conversioni: 0, ricavi: 0, valoreContratto: 0, incassato: 0, frequenza: 0,
            note: `TITOLO: ${c.titolo}\nCREATIVITÀ: ${c.creativo}\n\n${c.corpo}`,
          })), ...l,
        ]),
      });
    };

    return (
      <>
        <section className="card">
          <header>
            <h2>Piano di messaggio</h2>
            <span className="sp eyebrow">{pieno ? "completo" : "da completare"}</span>
            <button className="btn" onClick={() => apri("messaggio")}>{pieno ? "Modifica" : "Compila"}</button>
          </header>
          {sso && (
            <div style={{ padding: "18px 18px 4px" }}>
              <span className="eyebrow">La frase che tiene insieme tutto</span>
              <div style={{ fontSize: 19, fontWeight: 650, letterSpacing: "-.015em", marginTop: 6, maxWidth: "60ch" }}>{sso}</div>
            </div>
          )}
          <div className="wrap">
            <table className="tbl"><tbody>
              <tr><td className="sub" style={{ width: 210 }}>Mercato</td><td>{M.mercato || <span className="neg">da compilare</span>}</td></tr>
              <tr><td className="sub">Consapevolezza</td><td>{M.consapevolezza ? et(CONSAPEVOLEZZA, M.consapevolezza) : "—"}</td></tr>
              <tr><td className="sub">Dolore o guadagno</td><td>{M.dolore || <span className="neg">da compilare</span>}</td></tr>
              <tr><td className="sub">Dolori di chi compra</td><td>{M.doloriCompratore || "—"}</td></tr>
              <tr><td className="sub">Dolori dei curiosi</td><td className="sub">{M.doloriCuriosi || "—"}</td></tr>
              <tr><td className="sub">Soluzione unica</td><td>{M.soluzione || <span className="neg">da compilare</span>}</td></tr>
              <tr><td className="sub">Benefici</td><td>{ben.length ? ben.join(" · ") : <span className="neg">da compilare</span>}</td></tr>
              <tr><td className="sub">Obiezioni</td><td>{obi.length ? obi.join(" · ") : "—"}</td></tr>
              <tr><td className="sub">Offerta</td><td>{M.offerta || <span className="neg">da compilare</span>}</td></tr>
              <tr><td className="sub">Cosa puoi dire solo tu</td><td>{M.usp || "—"}</td></tr>
            </tbody></table>
          </div>
          {M.consapevolezza && (
            <div className="nota">
              <b>Gancio adatto a questo livello:</b> {GANCI[M.consapevolezza]}
            </div>
          )}
        </section>

        <section className="card">
          <header>
            <h2>I quattro annunci</h2>
            <span className="sp eyebrow">uno per fase del percorso d'acquisto</span>
            <button className="btn ghost mini" onClick={creaAnnunci}>Creali in Advertising</button>
          </header>
          {!pieno && (
            <div className="nota" style={{ paddingTop: 14 }}>
              Compila il piano di messaggio e le parti tra parentesi si riempiono da sole. Puoi comunque usare la struttura così com'è.
            </div>
          )}
          <div className="acts">
            {core4.map((c) => (
              <div className="tkrow" key={c.stadio}>
                <div className="t">
                  <b>{c.tit}</b>
                  <div style={{ marginTop: 8, fontSize: 13.3, whiteSpace: "pre-wrap", maxWidth: "78ch" }}>{c.corpo}</div>
                  <div className="meta" style={{ marginTop: 8 }}>
                    <span className="pill">titolo: {c.titolo}</span>
                    <span className="pill">creatività: {c.creativo}</span>
                  </div>
                </div>
              </div>
            ))}
          </div>
          <div className="nota">
            Per la fase 5, quella dopo l'acquisto, non si fanno annunci: si lavora con assistenza, follow-up e richiesta di recensioni.
          </div>
        </section>
      </>
    );
  };

  /* ---------- file e procedure ---------- */
  /* ---------- File KPI → Ciclo di vendita (scheda File e Ciclo di vendita del cliente) ---------- */
  const leggiFileKpi = (file) => {
    if (!attivo) return;
    const r = new FileReader();
    r.onload = () => {
      try {
        const wb = eExcel(file.name)
          ? XLSX.read(new Uint8Array(r.result), { type: "array", cellDates: true })
          : XLSX.read(String(r.result), { type: "string" });
        const fogli = wb.SheetNames.map((n) => ({ nome: n, righe: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: "" }) }));
        const tabelle = trovaTabelleKpi(fogli);
        setKpiImport({ nome: file.name, tabelle, scelte: tabelle.length ? [tabelle[0].id] : [], cliente: attivo.id,
          anno: new Date().getFullYear(), errore: tabelle.length ? "" : "Nel file non ho trovato tabelle con i mesi in riga e le voci (spesa, lead, appuntamenti, clienti, fatturato) in colonna." });
      } catch (e) { setKpiImport({ nome: file.name, tabelle: [], scelte: [], cliente: attivo.id, anno: new Date().getFullYear(), errore: `Non riesco a leggere il file: ${e.message || e}` }); }
    };
    if (eExcel(file.name)) r.readAsArrayBuffer(file); else r.readAsText(file);
  };
  /* Somma per mese delle tabelle scelte (es. FUNNEL OT + BOOST REEL di un mese senza colonna totale) */
  const mesiKpi = (k) => {
    const per = new Map();
    k.tabelle.filter((t) => k.scelte.includes(t.id)).forEach((t) => t.mesi.forEach((x) => {
      const periodo = x.periodo || `${k.anno}-${String(x.m).padStart(2, "0")}`;
      const v = per.get(periodo) || { periodo, m: x.m, anno: x.anno || k.anno, valori: {}, da: [] };
      Object.entries(x.valori).forEach(([c, n]) => { v.valori[c] = (v.valori[c] || 0) + n; });
      v.da.push(t.gruppo);
      per.set(periodo, v);
    }));
    return [...per.values()].sort((a, b) => (a.periodo < b.periodo ? -1 : 1));
  };
  const PannelloKpi = () => {
    if (!attivo) return null;
    const k = kpiImport;
    const mesi = k ? mesiKpi(k) : [];
    const clienti = [attivo.id, ...discendenti(attivo.id)].map(cliente).filter(Boolean);
    const campi = VOCI_KPI.filter((v) => mesi.some((x) => x.valori[v.k] !== undefined));
    const totaleEAltri = k && k.tabelle.some((t) => k.scelte.includes(t.id) && /totale/i.test(t.gruppo))
      && k.tabelle.some((t) => k.scelte.includes(t.id) && !/totale/i.test(t.gruppo) && t.foglio === k.tabelle.find((x) => k.scelte.includes(x.id) && /totale/i.test(x.gruppo)).foglio);
    const gia = k ? mesi.filter((x) => data.vendite.some((v) => v.clienteId === k.cliente && v.periodo === x.periodo)).length : 0;
    const importa = () => {
      const fonte = k.nome;
      setData((d) => {
        const r = applicaKpi(d.vendite, k.cliente, mesi.map((x) => ({ periodo: x.periodo, m: x.m, anno: x.anno, valori: x.valori })), { uid, fonte });
        return { ...d, vendite: r.vendite };
      });
      setKpiImport({ ...k, fatto: `Caricati ${mesi.length} mesi nel Ciclo di vendita di ${(cliente(k.cliente) || {}).nome || "—"}${gia ? ` (${gia} già presenti, aggiornati)` : ""}.` });
    };
    return conValuta(valutaCliente(data, (k && k.cliente) || attivo.id), () => (
      <section className="card" id="kpi-import">
        <header>
          <h2>Importa i numeri da un file KPI</h2>
          <span className="sp eyebrow">Excel o CSV · va nel Ciclo di vendita</span>
          {k && <button className="btn ghost mini" onClick={() => setKpiImport(null)}>Chiudi</button>}
        </header>
        <div style={{ padding: "14px 18px", display: "grid", gap: 10 }}>
          <div className="drop">
            <input type="file" accept=".xlsx,.xlsm,.xls,.ods,.csv" aria-label="File KPI"
              onChange={(e) => { const f = e.target.files && e.target.files[0]; if (f) leggiFileKpi(f); e.target.value = ""; }} />
            <span className="sub">Il foglio di riepilogo del cliente: voci in colonna (spesa, lead, appuntamenti fissati e svolti, nuovi clienti, fatturato, incassato) e mesi in riga, anche con una colonna per funnel sotto ogni mese.</span>
          </div>
          {k && k.errore && <div className="meta-msg" data-tipo="errore">{k.errore}</div>}
          {k && k.fatto && <div className="meta-msg" data-tipo="ok">{k.fatto}</div>}
          {k && k.tabelle.length > 0 && !k.fatto && (
            <>
              <div className="sub">Nel file <b>{k.nome}</b> ho trovato {k.tabelle.length} {k.tabelle.length === 1 ? "tabella" : "tabelle"}. Scegli quali caricare: se ne scegli più di una, i loro numeri si sommano mese per mese.</div>
              <div className="meta-camp" style={{ maxHeight: 240 }}>
                {k.tabelle.map((t) => (
                  <label key={t.id}>
                    <input type="checkbox" checked={k.scelte.includes(t.id)}
                      onChange={(e) => setKpiImport((x) => ({ ...x, scelte: e.target.checked ? [...x.scelte, t.id] : x.scelte.filter((y) => y !== t.id) }))} />
                    <span className="nome"><b>{t.gruppo}</b><span className="sub"> · foglio {t.foglio}</span></span>
                    <span className="sub mono">{t.mesi.length} mesi · {t.mesi[0].periodo ? meseIt(t.mesi[0].periodo) : t.mesi[0].m}{t.mesi.length > 1 ? ` – ${t.mesi[t.mesi.length - 1].periodo ? meseIt(t.mesi[t.mesi.length - 1].periodo) : t.mesi[t.mesi.length - 1].m}` : ""}</span>
                  </label>
                ))}
              </div>
              {totaleEAltri && <div className="meta-msg" data-tipo="errore">Hai scelto una colonna TOTALE insieme ad altre dello stesso foglio: i numeri verrebbero contati due volte.</div>}
              <div className="meta-riga">
                <span className="sub">Carica in</span>
                <select value={k.cliente} onChange={(e) => setKpiImport((x) => ({ ...x, cliente: e.target.value }))} aria-label="Cliente di destinazione" style={{ maxWidth: 300 }}>
                  {clienti.map((z) => <option key={z.id} value={z.id}>{z.id === attivo.id ? z.nome : `↳ ${z.nome}`}</option>)}
                </select>
                {k.tabelle.some((t) => k.scelte.includes(t.id) && t.annoMancante) && (
                  <>
                    <span className="sub">Anno dei mesi senza anno</span>
                    <input type="number" value={k.anno} min="2015" max="2100" style={{ width: 90 }}
                      onChange={(e) => setKpiImport((x) => ({ ...x, anno: Number(e.target.value) || x.anno }))} aria-label="Anno" />
                  </>
                )}
              </div>
              {mesi.length > 0 && (
                <div className="wrap"><table className="tbl">
                  <thead><tr><th>Mese</th>{campi.map((v) => <th key={v.k} className="r" title={v.l}>{v.b}</th>)}</tr></thead>
                  <tbody>
                    {mesi.map((x) => (
                      <tr key={x.periodo}>
                        <td><b>{meseIt(x.periodo)}</b>{data.vendite.some((v) => v.clienteId === k.cliente && v.periodo === x.periodo) ? <div className="sub">già presente: si aggiorna</div> : null}</td>
                        {campi.map((v) => <td key={v.k} className="r">{x.valori[v.k] === undefined ? "—" : ["spesa", "valoreContratto", "incassato"].includes(v.k) ? eur(x.valori[v.k]) : int(x.valori[v.k])}</td>)}
                      </tr>
                    ))}
                  </tbody>
                </table></div>
              )}
              <div className="meta-riga">
                <button className="btn ghost mini" disabled={!mesi.length} onClick={importa}>Carica {mesi.length} {mesi.length === 1 ? "mese" : "mesi"} nel Ciclo di vendita</button>
                <span className="sub">Appuntamenti fissati e svolti vanno in demo fissate e svolte, i nuovi clienti in chiusure. Si aggiornano solo i campi presenti nel file: gli altri restano come sono.</span>
              </div>
            </>
          )}
        </div>
      </section>
    ));
  };

  const SezioneFile = () => {
    const base = (attivo ? S.file : data.file)
      .filter((f) => (!fFile.cat || f.categoria === fFile.cat) && (!fFile.tipo || f.tipo === fFile.tipo))
      .filter((f) => cerca(`${f.titolo} ${f.tag} ${f.testo}`))
      .sort((a, b) => (a.aggiornato < b.aggiornato ? 1 : -1));
    const generali = attivo ? data.file.filter((f) => !f.clienteId) : [];
    const aperto = base.find((f) => f.id === fileAperto) || generali.find((f) => f.id === fileAperto);

    const Riga = ({ f, ombra }) => (
      <tr className="clic" onClick={() => setFileAperto(fileAperto === f.id ? "" : f.id)}>
        <td>
          <b>{f.titolo}</b>
          <div className="sub">
            {f.tipo === "documento" ? (f.testo || "").split("\n")[0].slice(0, 80) : f.url.replace(/^https?:\/\//, "").slice(0, 60)}
          </div>
        </td>
        <td><span className="tag" data-k={f.categoria === "procedura" ? "attivo" : "qualificato"}>{et(CATEGORIE_FILE, f.categoria)}</span></td>
        {!attivo && <td className="sub">{f.clienteId ? nomeCli(f.clienteId) : "Agenzia"}</td>}
        {ombra && <td className="sub">agenzia</td>}
        <td className="sub">{f.tag || "—"}</td>
        <td className="r sub">{dataIt(f.aggiornato)}</td>
        <td className="r">
          {f.tipo === "link"
            ? <button className="btn ghost mini" onClick={(e) => { e.stopPropagation(); window.open(f.url, "_blank", "noopener"); }}>Apri</button>
            : <button className="btn ghost mini" onClick={(e) => { e.stopPropagation(); apri("file", f); }}>Modifica</button>}
        </td>
      </tr>
    );

    const senzaProcedure = !data.file.some((f) => f.categoria === "procedura");

    return (
      <>
        {PannelloKpi()}
        {senzaProcedure && (
          <article className="al" data-g="attenzione">
            <div className="meta"><span className="tag">contenuti pronti</span></div>
            <h3>Tre procedure già scritte, da adattare</h3>
            <dl>
              <dt>Cosa sono</dt>
              <dd>Rilevazione settimanale, chiusura del mese e primi sette giorni con un nuovo cliente: i passaggi in ordine, con dentro le cose che se salti ti si ritorcono contro dopo due mesi.</dd>
              <dt>Cosa fare</dt>
              <dd>Le inserisco come documenti dell'agenzia, quindi visibili da ogni cliente. Poi le modifichi come vuoi: da quel momento sono tue.</dd>
            </dl>
            <div style={{ marginTop: 12 }}>
              <button className="btn" onClick={() => upd("file", (l) => [...PROCEDURE_BASE(), ...l])}>Inserisci le tre procedure</button>
            </div>
          </article>
        )}

        <section className="card">
          <header>
            <div className="ctrl">
              <select value={fFile.cat} onChange={(e) => setFFile({ ...fFile, cat: e.target.value })} aria-label="Filtra per categoria">
                <option value="">Tutte le categorie</option>
                {CATEGORIE_FILE.map((c) => <option key={c.v} value={c.v}>{c.l}</option>)}
              </select>
              <select value={fFile.tipo} onChange={(e) => setFFile({ ...fFile, tipo: e.target.value })} aria-label="Filtra per tipo">
                <option value="">Documenti e collegamenti</option>
                {TIPI_FILE.map((t) => <option key={t.v} value={t.v}>{t.l}</option>)}
              </select>
            </div>
            <span className="sp eyebrow">{base.length} elementi</span>
          </header>
          <div className="wrap">
            <table className="tbl">
              <thead><tr>
                <th>Titolo</th><th>Categoria</th>{!attivo && <th>Cliente</th>}<th>Etichette</th><th className="r">Aggiornato</th><th className="r">Azioni</th>
              </tr></thead>
              <tbody>{base.map((f) => <Riga key={f.id} f={f} />)}</tbody>
            </table>
          </div>
          {!base.length && (
            <div className="empty">
              <b>Niente qui</b>
              Metti i collegamenti alle cartelle Drive del cliente e scrivi le procedure che ripeti ogni mese.
            </div>
          )}
        </section>

        {aperto && aperto.tipo === "documento" && (
          <section className="card">
            <header>
              <h2>{aperto.titolo}</h2>
              <span className="sp eyebrow">aggiornato il {dataIt(aperto.aggiornato)}</span>
              <button className="btn ghost mini" onClick={() => apri("file", aperto)}>Modifica</button>
              <button className="btn ghost mini" onClick={() => setFileAperto("")}>Chiudi</button>
            </header>
            <div style={{ padding: "16px 18px", whiteSpace: "pre-wrap", fontSize: 13.6, lineHeight: 1.6, maxWidth: "80ch" }}>
              {aperto.testo || "Documento vuoto."}
            </div>
          </section>
        )}

        {attivo && generali.length > 0 && (
          <section className="card">
            <header><h2>Materiale dell'agenzia</h2><span className="sp eyebrow">valido per tutti i clienti</span></header>
            <div className="wrap">
              <table className="tbl">
                <thead><tr><th>Titolo</th><th>Categoria</th><th /><th>Etichette</th><th className="r">Aggiornato</th><th className="r">Azioni</th></tr></thead>
                <tbody>{generali.map((f) => <Riga key={f.id} f={f} ombra />)}</tbody>
              </table>
            </div>
          </section>
        )}

        <section className="card">
          <header><h2>Come conviene usarla</h2></header>
          <div style={{ padding: "14px 16px", fontSize: 13.3, color: "var(--ink2)" }}>
            I file veri restano dove sono già: Drive, Dropbox, Notion. Qui salvi il collegamento, così accanto ai numeri del cliente
            trovi anche la sua strategia, il brief creativo e il contratto, senza cercarli in tre posti.
            Le procedure invece conviene scriverle qui dentro: sono corte, cambiano spesso, e finiscono nel backup insieme a tutto il resto.
            <br /><br />
            Quello che l'app non fa è ospitare i file veri: un PDF o un video dentro l'archivio del browser lo riempirebbero in fretta,
            e ogni backup diventerebbe pesantissimo. Se ti serve un allegato accessibile, caricalo su Drive e incolla qui il collegamento.
          </div>
        </section>
      </>
    );
  };

  /* ---------- consulente ---------- */
  const SezioneAI = () => {
    /* Quattro ambiti: nessun dato (domanda generica), un cliente scelto,
       tutto il portafoglio, oppure l'agenzia come insieme. */
    const ambitoChat = chat.ambito || (attivo ? "cliente" : "agenzia");
    const cliScelto = data.clienti.find((z) => z.id === (chat.cliente || (attivo ? attivo.id : ""))) || attivo || null;
    const attiviChat = data.clienti.filter((z) => z.stato !== "chiuso" && z.archiviato !== "si");
    const cls = ambitoChat === "libera" ? []
      : ambitoChat === "cliente" ? (cliScelto ? [cliScelto] : [])
        : attiviChat;

    /* Il consulente può proporre automazioni già compilate: le descrive in JSON,
       l'app le mostra e tu decidi se crearle. Nessuna regola nasce senza il tuo clic. */
    const schemaAutomazioni = () => {
      const r = [];
      r.push("Puoi proporre automazioni da creare nel CRM. Se l'utente lo chiede, rispondi PRIMA con una spiegazione in italiano e POI con un blocco ```json contenente un array di regole, ognuna con questi campi:");
      r.push('{"nome": "testo", "trigger": "avviso|metrica|budget|task|chiusura|periodo", "clienteId": "" per tutti i clienti oppure l\'id indicato, "attiva": "si",');
      r.push('"filtroAvviso": "nome della metrica dell\'avviso, solo per trigger avviso", "soloCritici": "si|no",');
      r.push('"metrica": "ctr|cpc|cpm|cpa|roas|freq|sp|cv", "operatore": "gt|lt", "valore": numero, "giorno": numero per trigger periodo,');
      r.push('"azTask": "si|no", "taskTitolo": "testo", "taskPriorita": "urgente|alta|media|bassa", "taskGiorni": numero, "azWebhook": "no", "url": ""}');
      r.push("Le metriche del CRM: ctr è il CTR sul link in percentuale, cpa il costo per risultato, roas il ritorno in volte, freq la frequenza, sp la spesa totale, cv le conversioni.");
      r.push("Non inventare campi diversi da questi. Proponi poche regole ben mirate invece di molte generiche.");
      return r.join("\n");
    };

    const estraiRegole = (testo) => {
      const m = String(testo || "").match(/```json\s*([\s\S]*?)```/);
      if (!m) return null;
      try {
        const v = JSON.parse(m[1]);
        const lista = Array.isArray(v) ? v : [v];
        const trig = TRIGGER.map((x) => x.v);
        /* Il cliente indicato deve esistere davvero: un id inventato farebbe una regola che non scatta mai.
           Provo anche col nome; se non lo trovo la regola nasce spenta e lo dice nel nome. */
        const trovaCli = (v) => {
          if (!v) return "";
          const s = String(v).trim().toLowerCase();
          const z = data.clienti.find((c) => c.id === v) || data.clienti.find((c) => c.nome.trim().toLowerCase() === s);
          return z ? z.id : null;
        };
        const aree = data.aree.map((x) => x.id);
        return lista.filter((a) => a && a.nome && trig.includes(a.trigger)).map((a) => {
          const cid = trovaCli(a.clienteId);
          return {
            ...VUOTO.automazioni, ...a, id: uid(),
            clienteId: cid || "",
            nome: cid === null ? `[controlla il cliente] ${a.nome}` : a.nome,
            attiva: cid === null || a.attiva === "no" ? "no" : "si",
            taskArea: aree.includes(a.taskArea) ? a.taskArea : VUOTO.automazioni.taskArea,
            azWebhook: "no", url: "",
          };
        });
      } catch (e) { return null; }
    };

    const creaRegole = (regole) => {
      upd("automazioni", (l) => [...regole, ...l]);
      setChat((c) => ({ ...c, messaggi: [...c.messaggi, { da: "ai",
        testo: `Ho creato ${regole.length === 1 ? "la regola" : `${regole.length} regole`} in Automazioni: ${regole.map((r) => `"${r.nome}"`).join(", ")}. ${regole.some((r) => r.attiva === "no") ? "Quelle con un cliente che non ho trovato sono nate spente: controllale prima di attivarle. " : ""}Ricorda che le automazioni girano solo mentre l'app è aperta.` }] }));
    };

    /* Il contesto è la differenza fra una chat qualunque e un consulente che sa di cosa parla */
    const costruisciContesto = () => {
      const r = [];
      r.push(`Data di oggi: ${oggi()}. Agenzia: ${data.agenzia.nome}.`);
      if (ambitoChat === "libera") {
        r.push("L'utente ha scelto di non condividere i dati del CRM: rispondi alla domanda con le tue conoscenze, senza inventare numeri suoi.");
        return r.join("\n");
      }
      if (ambitoChat === "agenzia") {
        /* Numeri d'insieme: la domanda riguarda l'agenzia, non il singolo cliente */
        const mesi = righeMensili(attiviChat.filter((z) => !z.padreId).map((z) => z.id));
        const u = mesi[mesi.length - 1] || {};
        const attive = m.campagne.filter((k) => ["attiva", "test"].includes(k.stato));
        const crit = m.avvisi.filter((a) => a.gravita === "critica").length;
        r.push(`\n### AGENZIA NEL SUO INSIEME`);
        r.push(`${attiviChat.filter((z) => !z.padreId).length} clienti principali, ${attiviChat.filter((z) => z.padreId).length} sottoclienti, ${attive.length} campagne attive.`);
        if (u.mese) r.push(`Ultimo mese con dati (${u.mese}): spesa ${fmt(u.spesa)}, lead ${int(u.lead)}, CPL ${u.lead ? fmt(u.spesa / u.lead) : "n.d."}, chiusure ${int(u.chiusure)}, contratti ${fmt(u.contratto)}, incassato ${fmt(u.incassato)}.`);
        r.push(`Avvisi aperti ${m.avvisi.length}, di cui critici ${crit}.`);
        r.push(`Compensi ricorrenti: ${fmt(attiviChat.reduce((a, z) => a + (z.fee || 0), 0))} al mese.`);
        const perSem = ["verde", "giallo", "rosso", "grigio"].map((g) => `${g} ${attive.filter((k) => semaforo(data, k).segnale === g).length}`).join(", ");
        r.push(`Semaforo sulle attive: ${perSem}.`);
      }
      cls.forEach((z) => {
        /* Gli importi al consulente vanno con la valuta del cliente, non sempre in euro */
        const vz = valutaCliente(data, z.id);
        const camp = m.campagne.filter((k) => k.clienteId === z.id);
        const vend = data.vendite.filter((v) => v.clienteId === z.id).map(conChiusure);
        const chius = data.chiusure.filter((c) => c.clienteId === z.id);
        const tot = sommaVendite(vend);
        const rp = roasPareggio(z);
        r.push(`\n### CLIENTE: ${z.nome} (id "${z.id}", ${z.settore || "settore non indicato"}, stato ${z.stato})`);
        r.push(`Compenso ${z.fee || 0} ${valutaCompenso(data, z.id)}/mese, budget adv ${z.budgetMese || 0} ${vz}/mese, margine lordo ${z.margineLordo || 0}%, resi ${z.resi || 0}%${rp ? `, ROAS di pareggio ${fmt(rp)}` : ""}.`);
        if (z.note) r.push(`Note: ${z.note}`);
        if (camp.length) {
          r.push("CAMPAGNE (nome | piattaforma | obiettivo | stato | speso/budget | CTR | CPC | CPA | ROAS | frequenza):");
          camp.slice(0, 12).forEach((k) => r.push(`- ${k.nome} | ${et(PIATTAFORME, k.piattaforma)} | ${et(OBIETTIVI, k.obiettivo)} | ${k.stato} | ${fmt(k.d.sp)}/${fmt(k.budget)} ${vz} | ${fmt(k.d.ctr)}% | ${fmt(k.d.cpc)} ${vz} | ${k.d.cv ? fmt(k.d.cpa) + " " + vz : "n.d."} | ${k.d.rc ? fmt(k.d.roas) + "x" : "n.d."} | ${fmt(k.d.freq)}`));
          const al = camp.flatMap((k) => k.al.map((a) => `- [${a.gravita}] ${k.nome}: ${a.titolo} (${a.metrica} ${a.valore}, obiettivo ${a.soglia})`));
          if (al.length) r.push("AVVISI ATTIVI:\n" + al.join("\n"));
        }
        const cre = data.creativi.filter((c) => camp.some((k) => k.id === c.campagnaId)).map((c) => ({ ...c, d: calc(c) })).sort((a, b) => b.d.sp - a.d.sp).slice(0, 6);
        if (cre.length) {
          r.push("AD SET E ANNUNCI (nome | speso | CTR | CPA | frequenza):");
          cre.forEach((c) => r.push(`- ${c.nome} | ${fmt(c.d.sp)} ${vz} | ${fmt(c.d.ctr)}% | ${c.d.cv ? fmt(c.d.cpa) + " " + vz : "n.d."} | ${fmt(c.d.freq)}`));
        }
        if (vend.length) {
          r.push(`CICLO DI VENDITA (totali sui periodi caricati): lead ${tot.lead}, chiamate ${tot.chiamate}, discovery fissate ${tot.discoveryFissate} e svolte ${tot.discoverySvolte}, demo fissate ${tot.demoFissate} e svolte ${tot.demoSvolte}, chiusure ${tot.chiusure}, contratto ${fmt(tot.valoreContratto)} ${vz}, incassato ${fmt(tot.incassato)} ${vz}, spesa ${fmt(tot.spesa)} ${vz}.`);
          if (tot.slot) r.push(`Capacità: ${tot.slot} slot disponibili; tempo medio al primo contatto ${fmt(tot.oreContatto / vend.length, 1)} ore.`);
        }
        if (chius.length) {
          const perF = {};
          chius.forEach((c) => { perF[et(FUNNEL_PROV, c.funnel)] = (perF[et(FUNNEL_PROV, c.funnel)] || 0) + (c.valoreContratto || 0); });
          r.push(`CHIUSURE: ${chius.length}, contratto per funnel: ${Object.entries(perF).map(([k, v]) => `${k} ${fmt(v)} ${vz}`).join(", ")}.`);
        }
        const fn = data.funnel.filter((f) => f.clienteId === z.id);
        fn.forEach((f) => {
          r.push(`FUNNEL "${f.nome}": ` + (f.passi || []).map((p) => `${p.nome} ${int(p.visite)}→${int(p.conversioni)}${p.visite ? ` (${fmt((p.conversioni / p.visite) * 100, 1)}%)` : ""}`).join(" | "));
        });
        const tk = data.task.filter((t) => t.clienteId === z.id && t.stato !== "fatto").slice(0, 8);
        if (tk.length) r.push("TASK APERTI:\n" + tk.map((t) => `- ${t.titolo} (${t.priorita}, entro ${t.scadenza || "senza data"})`).join("\n"));
        const an = data.annotazioni.filter((a) => a.clienteId === z.id).slice(0, 5);
        if (an.length) r.push("DIARIO RECENTE:\n" + an.map((a) => `- ${a.data}: ${a.testo}`).join("\n"));
      });
      return r.join("\n");
    };

    const PREAMBOLO = `Sei un direttore marketing e media buyer esperto, e stai assistendo un collega che gestisce campagne per i clienti elencati qui sotto. Rispondi in italiano, in modo diretto e concreto, senza elenchi puntati inutili e senza ripetere i numeri che ti ha già dato se non servono al ragionamento. Basati solo sui dati forniti: se un dato manca, dillo e spiega quale servirebbe invece di stimarlo. Quando proponi azioni, mettile in ordine di impatto e sii specifico su cosa toccare. Tieni le risposte sotto le 250 parole salvo richiesta diversa.

Segui questo framework decisionale (sei soglie sequenziali, da The Ad Decision Framework di Jason Hornung), che è il riferimento scelto da chi ti sta chiedendo:
1. Non giudicare un annuncio prima che abbia speso due volte il valore medio dell'ordine, o due volte il costo per lead obiettivo. Sotto quella soglia si annota la tendenza e non si agisce.
2. CTR sul link sotto lo 0,8% dopo la soglia: problema di creatività, non di pubblico. Si spegne la creatività, si tiene il pubblico, si riscrive il gancio con un angolo diverso e la stessa offerta.
3. CPM oltre il 40% sopra la media dell'account: l'asta sta punendo la creatività. Si sostituisce l'annuncio dentro lo stesso ad set.
4. Sotto il ROAS obiettivo o sopra il CPL obiettivo dopo la soglia, senza miglioramento negli ultimi tre giorni: si spegne e si documenta l'angolo fallito, senza ritestare lo stesso gancio sullo stesso pubblico. Un solo periodo negativo invece è rumore.
5. Slope test: se era a obiettivo al lancio e peggiora da tre o più settimane consecutive senza cause esterne, è affaticamento lento. Si limita il budget e si lancia subito una creatività sfidante, senza aspettare il fondo.
6. Eccezione: frequenza sopra 4,5 ma risultati ancora a obiettivo significa che la creatività tiene. Non si spegne per abitudine.
Aggiungi sempre la lettura economica: un ROAS ha senso solo confrontato con il ROAS di pareggio ricavato dal margine del cliente.

Tieni presente anche il metodo in sette passi (AOA Way, stesso autore), perché spesso il problema non è nella campagna ma a monte: 1) piano di messaggio che copre tutte le fasi del processo d'acquisto — riconoscimento del problema, raccolta di informazioni, valutazione delle soluzioni, acquisto — e non solo la prima e la quarta, che è la causa più comune del logoramento rapido; 2) piano di conversione scelto come sistema unico; 3) asset pronti e revisionati; 4) dati attendibili, con scarto noto fra piattaforme e numeri reali dell'azienda; 5) test di validazione prima di scalare; 6) regole operative su spesa, scala e spegnimento; 7) ottimizzazione mirata sul solo punto di fallimento del percorso. Quando i dati mostrano campagne che si consumano in fretta o traffico che non converte, verifica prima quale di questi passi manca, invece di suggerire ritocchi al targeting.

Sul messaggio, il metodo prevede quattro annunci, uno per fase: freddo con dolore, soluzione e tre benefici; tiepido con la guida all'acquisto ("leggi questo prima di acquistare X"); tiepido con il chi siamo che dichiara la missione e supera l'obiezione principale; caldo con l'offerta e la proiezione nel futuro. Per la fase 5, dopo l'acquisto, non si fanno annunci. Ricorda inoltre due distinzioni: i dolori di chi compra sono attuali e urgenti, quelli dei curiosi sono veri ma non urgenti, e costruire il messaggio sui secondi riempie il calendario di persone che non comprano; il gancio va scelto in base a quanto il mercato conosce già problema e soluzioni. Se vedi show up basso o qualità dei lead scarsa, la prima ipotesi da verificare è il messaggio corto, cioè l'assenza degli annunci delle fasi 2 e 3, non il lavoro del commerciale.

${schemaAutomazioni()}

DATI DISPONIBILI:
`;

    const invia = async (testo) => {
      const domanda = (testo || chat.testo).trim();
      if (!domanda || chat.attesa) return;
      const storico = [...chat.messaggi, { da: "io", testo: domanda }];
      setChat((c) => ({ ...c, messaggi: storico, testo: "", attesa: true, errore: "" }));
      try {
        /* Il contesto va nel messaggio di sistema e si ricostruisce a ogni domanda:
           prima partiva solo col primo messaggio e le domande successive lavoravano su numeri vecchi. */
        const messages = [];
        storico.filter((x) => x.da !== "errore").forEach((x) => {
          const role = x.da === "io" ? "user" : "assistant";
          const ultimo = messages[messages.length - 1];
          if (ultimo && ultimo.role === role) ultimo.content += "\n\n" + x.testo;
          else messages.push({ role, content: x.testo });
        });
        while (messages.length && messages[0].role !== "user") messages.shift();
        const intestazioni = { "Content-Type": "application/json" };
        if (data.aiKey) { intestazioni["x-api-key"] = data.aiKey; intestazioni["anthropic-version"] = "2023-06-01"; intestazioni["anthropic-dangerous-direct-browser-access"] = "true"; }
        const risposta = await fetch("https://api.anthropic.com/v1/messages", {
          method: "POST", headers: intestazioni,
          body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: 4000, system: PREAMBOLO + costruisciContesto(), messages }),
        });
        const j = await risposta.json();
        if (j.error) throw new Error(j.error.message || "errore");
        const testoAI = (j.content || []).filter((x) => x.type === "text").map((x) => x.text).join("\n").trim();
        setChat((c) => ({ ...c, messaggi: [...storico, { da: "ai", testo: testoAI || "Non ho ricevuto risposta." }], attesa: false }));
      } catch (e) {
        console.error(e);
        setChat((c) => ({
          ...c, attesa: false,
          messaggi: [...storico, { da: "errore", testo: "Non sono riuscito a rispondere. Se stai usando l'app fuori da Claude serve una chiave API personale, che puoi incollare qui sotto." }],
        }));
      }
    };

    const inTask = (t) => {
      const righe = t.split("\n").filter(Boolean);
      apri("task", { titolo: (righe[0] || "").slice(0, 90), note: t }, true);
    };
    const inDiario = (t) => {
      upd("annotazioni", (l) => [{ id: uid(), clienteId: attivo ? attivo.id : "", data: oggi(), tipo: "altro", campagnaId: "", testo: t }, ...l]);
    };

    /* --- lettura e analisi di un report allegato --- */
    const leggiAllegato = (file) => {
      const testuale = /\.(csv|tsv|txt|json|md)$/i.test(file.name) || (file.type || "").startsWith("text");
      const r = new FileReader();
      r.onload = () => {
        if (eExcel(file.name)) {
          try {
            const fogli = fogliDaExcel(r.result);
            if (!fogli.length) { setEstrazione({ errore: "Il foglio di calcolo non contiene righe leggibili." }); return; }
            const testo = fogli.map((f) => `### Foglio: ${f.nome}\n${f.csv}`).join("\n\n").slice(0, 60000);
            setAllegato({ nome: `${file.name} · ${fogli.length} ${fogli.length === 1 ? "foglio" : "fogli"}`, tipo: "testo", testo });
          } catch (e) {
            setEstrazione({ errore: "Non sono riuscito ad aprire il foglio di calcolo: prova a salvarlo in CSV." });
          }
          return;
        }
        if (testuale) setAllegato({ nome: file.name, tipo: "testo", testo: String(r.result).slice(0, 60000) });
        else {
          const b64 = String(r.result).split(",")[1] || "";
          setAllegato({ nome: file.name, tipo: file.type === "application/pdf" ? "pdf" : "immagine", dati: b64, media: file.type || "image/png" });
        }
      };
      r.onerror = () => setEstrazione({ errore: "Non sono riuscito a leggere il file." });
      if (eExcel(file.name)) r.readAsArrayBuffer(file);
      else if (testuale) r.readAsText(file);
      else r.readAsDataURL(file);
    };

    const SCHEMA = `{
  "campagne": [{"nome":"", "piattaforma":"meta|google|tiktok|linkedin|youtube|amazon|altro", "obiettivo":"vendite|lead|traffico|awareness", "spesa":0, "impression":0, "click":0, "clickLink":0, "lead":0, "conversioni":0, "ricavi":0, "frequenza":0}],
  "creativi": [{"nome":"", "campagna":"", "tipo":"adset|annuncio", "spesa":0, "impression":0, "click":0, "clickLink":0, "conversioni":0, "ricavi":0, "frequenza":0}],
  "vendite": [{"periodo":"AAAA-MM", "lead":0, "chiamate":0, "discoveryFissate":0, "discoverySvolte":0, "demoFissate":0, "demoSvolte":0, "offerte":0, "spesa":0, "slot":0, "oreContatto":0}],
  "chiusure": [{"nome":"", "azienda":"", "email":"", "telefono":"", "data":"AAAA-MM-GG", "valoreContratto":0, "incassato":0, "campagna":"", "adset":"", "annuncio":"", "funnel":"vsl|webinar|form|dm|chiamata|referral|organico|evento|altro", "closer":""}],
  "note": "una riga su cosa hai trovato e su cosa non eri sicuro"
}`;

    const analizza = async () => {
      const cid = chat.destinazione || (attivo ? attivo.id : "");
      if (!allegato || !cid) { setEstrazione({ errore: "Scegli il cliente di destinazione prima di analizzare." }); return; }
      setEstrazione({ attesa: true });
      const nomiCamp = data.campagne.filter((k) => k.clienteId === cid).map((k) => k.nome);
      const istruzioni = `Estrai i dati da questo report pubblicitario o commerciale e restituisci SOLO JSON valido, senza testo prima o dopo e senza blocchi markdown.

Schema da rispettare esattamente:
${SCHEMA}

Regole:
- I numeri devono essere numeri, con il punto come separatore decimale, senza simboli di valuta e senza separatore delle migliaia.
- Ometti le liste vuote e i campi che il report non contiene: non inventare valori e non stimarli.
- Le metriche derivate (CTR, CPC, CPM, CPA, ROAS) NON vanno riportate: si ricalcolano da sole.
- "periodo" e' il mese in formato AAAA-MM; "data" e' AAAA-MM-GG.
- Massimo 15 elementi per lista: se ce ne sono di piu', tieni quelli con la spesa piu' alta.
- Se riconosci una campagna gia' presente in questo elenco, usa esattamente lo stesso nome: ${nomiCamp.length ? nomiCamp.join(" | ") : "(nessuna campagna gia' inserita)"}.
- Se il file non contiene dati utilizzabili, restituisci solo {"note":"spiegazione"}.`;

      const contenuto = [];
      if (allegato.tipo === "pdf") contenuto.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: allegato.dati } });
      if (allegato.tipo === "immagine") contenuto.push({ type: "image", source: { type: "base64", media_type: allegato.media, data: allegato.dati } });
      contenuto.push({ type: "text", text: istruzioni + (allegato.tipo === "testo" ? `\n\nCONTENUTO DEL FILE ${allegato.nome}:\n${allegato.testo}` : "") });

      try {
        const intestazioni = { "Content-Type": "application/json" };
        if (data.aiKey) { intestazioni["x-api-key"] = data.aiKey; intestazioni["anthropic-version"] = "2023-06-01"; intestazioni["anthropic-dangerous-direct-browser-access"] = "true"; }
        const risposta = await fetch("https://api.anthropic.com/v1/messages", {
          method: "POST", headers: intestazioni,
          body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: 4000, messages: [{ role: "user", content: contenuto }] }),
        });
        const j = await risposta.json();
        if (j.error) throw new Error(j.error.message || "errore");
        const grezzo = (j.content || []).filter((x) => x.type === "text").map((x) => x.text).join("").trim()
          .replace(/^```(json)?/i, "").replace(/```$/, "").trim();
        let est;
        try { est = JSON.parse(grezzo); }
        catch (e2) { throw new Error("La risposta non e' arrivata in forma leggibile: prova con un file piu' piccolo o con meno righe."); }
        setEstrazione({ ...est, clienteId: cid });
      } catch (e) {
        setEstrazione({ errore: e.message || "Analisi non riuscita." });
      }
    };

    const conta = (e) => ["campagne", "creativi", "vendite", "chiusure"].reduce((s, k) => s + ((e[k] || []).length), 0);

    const applicaEstratti = () => {
      const e = estrazione, cid = e.clienteId;
      const norm = (s) => String(s || "").trim().toLowerCase();
      let creati = 0, aggiornati = 0;
      /* Calcolo sui dati attuali e poi salvo: con setData(funzione) i contatori
         venivano letti prima che la funzione girasse, e il messaggio diceva sempre 0. */
      const ndFinale = ((d) => {
        const nd = { ...d };
        (e.campagne || []).forEach((r) => {
          if (!r.nome) return;
          const met = {};
          ["spesa", "impression", "click", "clickLink", "lead", "conversioni", "ricavi", "frequenza"].forEach((k) => { if (r[k] !== undefined) met[k] = parseNumIntl(r[k]); });
          const i = nd.campagne.findIndex((k) => norm(k.nome) === norm(r.nome) && k.clienteId === cid);
          if (i > -1) { nd.campagne = nd.campagne.map((k, j) => (j === i ? { ...k, ...met } : k)); aggiornati++; }
          else {
            nd.campagne = [{
              id: uid(), clienteId: cid, nome: r.nome, piattaforma: r.piattaforma || "altro", obiettivo: r.obiettivo || "vendite",
              stato: "attiva", inizio: oggi(), fine: piu(30), budgetGiorno: 0, budget: 0, spesa: 0, impression: 0, click: 0, clickLink: 0, lead: 0, conversioni: 0, ricavi: 0,
              frequenza: 0, bmk: { ...BMK_DEFAULT[r.obiettivo || "vendite"] }, ...met,
            }, ...nd.campagne];
            creati++;
          }
        });
        (e.creativi || []).forEach((r) => {
          if (!r.nome) return;
          const camp = nd.campagne.find((k) => norm(k.nome) === norm(r.campagna) && k.clienteId === cid);
          const met = {};
          ["spesa", "impression", "click", "clickLink", "lead", "conversioni", "ricavi", "frequenza"].forEach((k) => { if (r[k] !== undefined) met[k] = parseNumIntl(r[k]); });
          const i = nd.creativi.findIndex((c) => norm(c.nome) === norm(r.nome) && c.clienteId === cid);
          if (i > -1) { nd.creativi = nd.creativi.map((c, j) => (j === i ? { ...c, ...met } : c)); aggiornati++; }
          else if (camp) {
            nd.creativi = [{ id: uid(), clienteId: cid, campagnaId: camp.id, tipo: r.tipo === "annuncio" ? "annuncio" : "adset", nome: r.nome, stato: "attiva", spesa: 0, impression: 0, click: 0, conversioni: 0, ricavi: 0, valoreContratto: 0, incassato: 0, frequenza: 0, note: "", ...met }, ...nd.creativi];
            creati++;
          }
        });
        (e.vendite || []).forEach((r) => {
          const periodo = normalizzaMese(r.periodo);
          if (!periodo) return;
          const campi = {};
          ["lead", "chiamate", "discoveryFissate", "discoverySvolte", "demoFissate", "demoSvolte", "offerte", "spesa", "slot", "oreContatto"]
            .forEach((k) => { if (r[k] !== undefined) campi[k] = parseNumIntl(r[k]); });
          const i = nd.vendite.findIndex((v) => v.clienteId === cid && v.periodo === periodo);
          if (i > -1) { nd.vendite = nd.vendite.map((v, j) => (j === i ? { ...v, ...campi } : v)); aggiornati++; }
          else {
            nd.vendite = [{ id: uid(), clienteId: cid, periodo, lead: 0, chiamate: 0, discoveryFissate: 0, discoverySvolte: 0, demoFissate: 0, demoSvolte: 0, offerte: 0, chiusure: 0, valoreContratto: 0, incassato: 0, spesa: 0, slot: 0, oreContatto: 0, ...campi }, ...nd.vendite];
            creati++;
          }
        });
        (e.chiusure || []).forEach((r) => {
          if (!r.nome) return;
          const dataC = normalizzaData(r.data) || oggi();
          if (nd.chiusure.some((c) => c.clienteId === cid && norm(c.nome) === norm(r.nome) && c.data === dataC)) return;
          const camp = nd.campagne.find((k) => norm(k.nome) === norm(r.campagna) && k.clienteId === cid);
          nd.chiusure = [{
            id: uid(), clienteId: cid, data: dataC, nome: r.nome, azienda: r.azienda || "", email: r.email || "", telefono: r.telefono || "",
            funnel: (FUNNEL_PROV.find((f) => f.v === r.funnel) || { v: "altro" }).v, campagnaId: camp ? camp.id : "",
            adset: r.adset || "", annuncio: r.annuncio || "", closer: r.closer || "",
            valoreContratto: parseNumIntl(r.valoreContratto), incassato: parseNumIntl(r.incassato), note: "",
          }, ...nd.chiusure];
          creati++;
        });
        return nd;
      })(data);
      setData(ndFinale);
      setEstrazione(null); setAllegato(null);
      setChat((c) => ({ ...c, messaggi: [...c.messaggi, { da: "ai", testo: `Ho inserito i dati del report: ${creati} elementi creati e ${aggiornati} aggiornati. Controllali in Advertising e in Ciclo di vendita prima di usarli con il cliente.` }] }));
    };

    const RIEPILOGO = [["campagne", "Campagne"], ["creativi", "Ad set e annunci"], ["vendite", "Periodi commerciali"], ["chiusure", "Chiusure"]];

    return (
      <>
        <section className="card">
          <header><h2>Analizza un report</h2><span className="sp eyebrow">Excel, CSV, PDF o schermata</span></header>
          <div style={{ padding: "16px 18px", display: "grid", gap: 12, maxWidth: 640 }}>
            <div className="ctrl">
              <select value={chat.destinazione || (attivo ? attivo.id : "")} onChange={(ev) => setChat({ ...chat, destinazione: ev.target.value })} aria-label="Cliente di destinazione">
                <option value="">— scegli il cliente —</option>
                {attivi.map((z) => <option key={z.id} value={z.id}>{z.nome}</option>)}
              </select>
            </div>
            <input type="file" accept=".csv,.tsv,.txt,.pdf,.xlsx,.xlsm,.xls,.ods,.png,.jpg,.jpeg,.webp"
              onChange={(ev) => { const fl = ev.target.files && ev.target.files[0]; if (fl) leggiAllegato(fl); ev.target.value = ""; }} />
            {allegato && (
              <div className="riga">
                <span className="tag" data-k="qualificato">{allegato.tipo}</span>
                <span className="sub">{allegato.nome}</span>
                <span style={{ marginLeft: "auto" }} />
                <button className="btn ghost mini" onClick={() => { setAllegato(null); setEstrazione(null); }}>Togli</button>
                <button className="btn" onClick={analizza} disabled={!!(estrazione && estrazione.attesa)}>
                  {estrazione && estrazione.attesa ? "Sto leggendo…" : "Analizza e proponi"}
                </button>
              </div>
            )}
            <span className="hint">Leggo il file, estraggo i numeri e ti mostro cosa scriverei. Niente viene salvato finché non confermi tu.</span>
          </div>
        </section>

        {estrazione && estrazione.errore && (
          <section className="card"><div className="empty"><b>Non ha funzionato</b>{estrazione.errore}</div></section>
        )}

        {estrazione && !estrazione.errore && !estrazione.attesa && (
          <section className="card">
            <header><h2>Cosa ho trovato</h2><span className="sp eyebrow">per {nomeCli(estrazione.clienteId)}</span></header>
            {estrazione.note && <div className="nota" style={{ paddingTop: 14 }}>{estrazione.note}</div>}
            {conta(estrazione) === 0 ? (
              <div className="empty"><b>Nessun dato utilizzabile</b>Il file non conteneva numeri che sapessi collocare.</div>
            ) : (
              <>
                {RIEPILOGO.filter(([k]) => (estrazione[k] || []).length).map(([k, l]) => (
                  <div key={k}>
                    <div className="gruppo">{l} · {estrazione[k].length}</div>
                    <div className="wrap">
                      <table className="tbl"><tbody>
                        {estrazione[k].map((r, i) => (
                          <tr key={i}>
                            <td><b>{r.nome || meseIt(normalizzaMese(r.periodo)) || "—"}</b>{r.campagna && <div className="sub">{r.campagna}</div>}</td>
                            <td className="sub" style={{ fontSize: 12 }}>
                              {Object.entries(r).filter(([c, v]) => !["nome", "campagna", "periodo", "tipo"].includes(c) && v !== "" && v !== undefined)
                                .map(([c, v]) => `${c}: ${typeof v === "number" ? fmt(v, 2) : v}`).join(" · ")}
                            </td>
                          </tr>
                        ))}
                      </tbody></table>
                    </div>
                  </div>
                ))}
                <div style={{ padding: 16, display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                  <button className="btn" onClick={applicaEstratti}>Inserisci nei dati</button>
                  <button className="btn ghost" onClick={() => setEstrazione(null)}>Scarta</button>
                  <span className="sub">Campagne e periodi già presenti vengono aggiornati, gli altri creati.</span>
                </div>
              </>
            )}
          </section>
        )}

        <section className="card">
          <header>
            <div className="ctrl">
              <select value={ambitoChat} onChange={(e) => setChat({ ...chat, ambito: e.target.value })} aria-label="Di cosa parliamo">
                <option value="agenzia">L'agenzia nel suo insieme</option>
                <option value="cliente">Un cliente in particolare</option>
                <option value="tutto">Tutti i clienti, uno per uno</option>
                <option value="libera">Domanda generica, senza i miei dati</option>
              </select>
              {ambitoChat === "cliente" && (
                <select value={(chat.cliente || (attivo ? attivo.id : ""))}
                  onChange={(e) => setChat({ ...chat, cliente: e.target.value })} aria-label="Quale cliente">
                  <option value="">— scegli —</option>
                  {ordinaGerarchia().map(({ z, liv }) => (
                    <option key={z.id} value={z.id}>{liv ? `${"\u00A0".repeat(liv * 4)}↳ ` : ""}{z.nome}</option>
                  ))}
                </select>
              )}
            </div>
            <span className="sp" />
            {chat.messaggi.length > 0 && (
              <button className="btn ghost mini" onClick={() => setChat({ ...chat, messaggi: [], errore: "" })}>Nuova conversazione</button>
            )}
          </header>

          <div className="chat">
            {!chat.messaggi.length && (
              <div className="empty">
                <b>{ambitoChat === "libera" ? "Chiedimi quello che vuoi" : "Chiedimi qualcosa sui tuoi numeri"}</b>
                {ambitoChat === "libera"
                  ? "In questa modalità non ti guardo i dati: rispondo con quello che so di advertising e metodo."
                  : ambitoChat === "agenzia"
                    ? `Ho davanti i numeri dell'agenzia: ${attiviChat.filter((z) => !z.padreId).length} clienti principali, le campagne attive, gli avvisi aperti e l'andamento dell'ultimo mese.`
                    : ambitoChat === "cliente"
                      ? (cliScelto
                        ? `Ho davanti campagne, avvisi, ciclo di vendita, chiusure, funnel e task di ${cliScelto.nome}.`
                        : "Scegli un cliente dal menu qui sopra.")
                      : "Ho davanti campagne, avvisi, ciclo di vendita, chiusure, funnel e task di tutti i clienti attivi."}
              </div>
            )}
            {!chat.messaggi.length && (
              <div className="spunti">
                <span className="eyebrow">Da cosa partire</span>
                {(ambitoChat === "agenzia" ? [
                  "Quali clienti stanno andando peggio questo mese e perché?",
                  "Dove sto perdendo più soldi in rapporto a quanto incasso?",
                  "Creami le automazioni che mi avvisano quando una campagna supera le soglie del framework",
                  "Come dovrei distribuire il mio tempo fra i clienti che ho?",
                ] : ambitoChat === "libera" ? [
                  "Spiegami la differenza fra CTR totale e CTR sul link",
                  "Come si struttura un'offerta per un avvocato che fa separazioni?",
                  "Quando conviene scalare in verticale invece che in orizzontale?",
                  "Che differenza c'è fra pubblico freddo, tiepido e caldo?",
                ] : [
                  "Cosa non va in questo cliente e da dove comincio?",
                  "Il messaggio è corto? Quali annunci mi mancano?",
                  "Che angolo proveresti per la prossima campagna?",
                  "Creami un'automazione che mi avvisa quando questo cliente va fuori soglia",
                ]).map((q, i2) => (
                  <button key={i2} className="btn ghost mini" onClick={() => setChat({ ...chat, testo: q })}>{q}</button>
                ))}
              </div>
            )}
            {chat.messaggi.map((x, i) => (
              <div className="msg" data-da={x.da === "io" ? "io" : "ai"} key={i}>
                {x.testo}
                {x.da === "ai" && (() => {
                  const regole = estraiRegole(x.testo);
                  return (
                    <>
                      {regole && regole.length > 0 && (
                        <div className="proposta">
                          <span className="eyebrow">
                            {regole.length === 1 ? "Automazione proposta" : `${regole.length} automazioni proposte`}
                          </span>
                          {regole.map((r, j) => (
                            <div className="riga" key={j} style={{ gap: 9, alignItems: "flex-start" }}>
                              <span className="tag">{et(TRIGGER, r.trigger)}</span>
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <b>{r.nome}</b>
                                <div className="sub">
                                  {r.clienteId ? nomeCli(r.clienteId) : "tutti i clienti"}
                                  {r.azTask === "si" ? ` · crea un task ${et(PRIORITA, r.taskPriorita)} a ${int(r.taskGiorni || 2)} giorni` : " · solo registro"}
                                </div>
                              </div>
                            </div>
                          ))}
                          <div className="riga" style={{ gap: 8 }}>
                            <button className="btn" onClick={() => creaRegole(regole)}>
                              {regole.length === 1 ? "Crea la regola" : `Crea le ${regole.length} regole`}
                            </button>
                            <button className="btn ghost mini" onClick={() => { setVista("automazioni"); }}>Vai alle automazioni</button>
                          </div>
                        </div>
                      )}
                      <div className="azioni">
                        <button className="btn ghost mini" onClick={() => inTask(x.testo)}>Trasforma in task</button>
                        <button className="btn ghost mini" onClick={() => inDiario(x.testo)}>Salva nel diario</button>
                      </div>
                    </>
                  );
                })()}
              </div>
            ))}
            {chat.attesa && (
              <div className="msg" data-da="ai"><span className="puntini"><i /><i /><i /></span></div>
            )}
          </div>

          <div className="compositore">
            <textarea rows={2} value={chat.testo} placeholder="Scrivi qui la tua domanda…"
              onChange={(e) => setChat({ ...chat, testo: e.target.value })}
              onKeyDown={(e) => { if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) invia(); }} />
            <button className="btn" onClick={() => invia()} disabled={chat.attesa}>{chat.attesa ? "…" : "Invia"}</button>
          </div>
        </section>

        <section className="card">
          <header><h2>Come funziona, e dove smette di funzionare</h2></header>
          <div style={{ padding: "14px 16px", fontSize: 13.3, color: "var(--ink2)" }}>
            A ogni domanda l'app allega un riepilogo dei tuoi dati: per questo le risposte parlano dei tuoi clienti e non in generale.
            Le risposte si possono trasformare in task o archiviare nel diario, così non restano in una chat che poi si perde.
            <br /><br />
            Dentro Claude non serve configurare nulla. Se un domani userai l'app come file autonomo sul tuo computer, questa sezione
            avrà bisogno di una chiave API personale: incollala qui sotto. Non metterla mai in una versione pubblicata online,
            perché chiunque apra la pagina potrebbe usarla a tue spese.
          </div>
          <div style={{ padding: "0 16px 16px", maxWidth: 480 }}>
            <div className="f">
              <label htmlFor="ai-key">Chiave API (solo per l'uso fuori da Claude)</label>
              <input id="ai-key" type="password" value={data.aiKey || ""} placeholder="sk-ant-…"
                onChange={(e) => setData((d) => ({ ...d, aiKey: e.target.value }))} />
            </div>
          </div>
        </section>
      </>
    );
  };

  /* ---------- guida ---------- */
  const SezioneGuida = () => {
    const cap = GUIDA.find((g) => g.id === capitolo) || GUIDA[0];
    const passi = [
      { l: "Logo e costo orario impostati", ok: !!(data.agenzia.logo || data.agenzia.costoOrario), dove: "Impostazioni" },
      { l: "Almeno un cliente", ok: attivi.filter((z) => !z.padreId).length > 0, dove: "Clienti" },
      { l: "Margine del cliente compilato", ok: data.clienti.some((z) => z.margineLordo > 0), dove: "Scheda cliente" },
      { l: "Campagne con benchmark", ok: data.campagne.some((k) => k.bmk && (k.bmk.roasMin || k.bmk.cplMax || k.bmk.cpaMax || k.bmk.ctrMin)), dove: "Advertising" },
      { l: "Valore medio ordine o CPL/CPA obiettivo", ok: data.clienti.some((z) => z.aov > 0) || data.campagne.some((k) => tettoCpl(k.bmk) > 0), dove: "Scheda cliente" },
      { l: "Budget giornaliero sulle campagne", ok: data.campagne.some((k) => k.budgetGiorno > 0), dove: "Advertising" },
      { l: "Fasi della campagna compilate", ok: (data.periodi || []).some((pz) => (pz.spesa || 0) > 0), dove: "Advertising · Fasi" },
      { l: "Piano di messaggio di un cliente", ok: data.clienti.some((z) => z.msg && z.msg.mercato), dove: "Messaggio" },
      { l: "KPI impostato sulle campagne attive", ok: m.attive.length > 0 && m.attive.every((k) => kpiCosto(k, k.d) > 0), dove: "Advertising" },
      { l: "Account pubblicitario sui clienti", ok: attivi.some((z) => (z.accountAdv || "").trim()), dove: "Scheda cliente" },
      { l: "Clic sul link inseriti", ok: data.campagne.some((k) => k.clickLink > 0) || (data.periodi || []).some((pz) => pz.clickLink > 0), dove: "Advertising" },
      { l: "Date di fine sulle campagne chiuse", ok: !data.campagne.some((k) => ["conclusa", "pausa"].includes(k.stato) && !k.fine), dove: "Advertising" },
      { l: "Prima rilevazione salvata", ok: data.rilevazioni.length > 0, dove: "Advertising" },
      { l: "Registro dei venditori caricato", ok: (data.giornate || []).length > 0, dove: "Ciclo di vendita" },
      { l: "Un periodo commerciale", ok: data.vendite.length > 0, dove: "Ciclo di vendita" },
      { l: "Una chiusura registrata", ok: data.chiusure.length > 0, dove: "Chiusure" },
      { l: "Backup scaricato", ok: !!data.backup.ultimoFile, dove: "Backup" },
      { l: "Procedure e documenti in archivio", ok: data.file.length > 0, dove: "File" },
    ];
    const fatti = passi.filter((p) => p.ok).length;

    return (
      <>
        <section className="card">
          <header><h2>A che punto sei</h2><span className="sp eyebrow">{fatti} di {passi.length} passaggi</span></header>
          <div className="barra" style={{ margin: "0 16px", height: 6 }}>
            <i style={{ width: `${(fatti / passi.length) * 100}%` }} />
          </div>
          <div className="wrap">
            <table className="tbl"><tbody>
              {passi.map((p) => (
                <tr key={p.l}>
                  <td style={{ width: 34 }}><span className="tag" data-k={p.ok ? "attivo" : ""}>{p.ok ? "✓" : "—"}</span></td>
                  <td><b style={{ fontWeight: p.ok ? 500 : 600, color: p.ok ? "var(--ink2)" : "var(--ink)" }}>{p.l}</b></td>
                  <td className="r sub">{p.dove}</td>
                </tr>
              ))}
            </tbody></table>
          </div>
        </section>

        <section className="card">
          <header>
            <div className="ctrl">
              {GUIDA.map((g) => (
                <button key={g.id} className="btn ghost mini" data-on={cap.id === g.id ? "1" : "0"} onClick={() => setCapitolo(g.id)}>{g.t}</button>
              ))}
            </div>
          </header>
          <div style={{ padding: "16px 18px" }}>
            <h2 style={{ margin: "0 0 6px", fontSize: 18, fontWeight: 650, letterSpacing: "-.015em" }}>{cap.t}</h2>
            <p style={{ margin: 0, color: "var(--ink2)", fontSize: 13.5, maxWidth: "70ch" }}>{cap.intro}</p>
          </div>
          <div className="acts">
            {cap.blocchi.map((b) => (
              <div className="tkrow" key={b.h}>
                <div className="t">
                  <b>{b.h}</b>
                  <div className="sub" style={{ fontSize: 13.2, marginTop: 3, maxWidth: "80ch" }}>{b.p}</div>
                </div>
              </div>
            ))}
          </div>
        </section>

        <section className="card">
          <header><h2>Glossario</h2><span className="sp eyebrow">cosa significano davvero i numeri</span></header>
          <div className="wrap">
            <table className="tbl"><tbody>
              {GLOSSARIO.map(([k, v]) => (
                <tr key={k}><td style={{ width: 180 }}><b>{k}</b></td><td className="sub" style={{ fontSize: 13 }}>{v}</td></tr>
              ))}
            </tbody></table>
          </div>
        </section>
      </>
    );
  };

  /* ---------- importazione ---------- */
  const SezioneImporta = () => {
    const cfg = CAMPI_IMPORT[imp.tipo];
    const pronti = imp.intestazioni.length > 0;

    const caricaTesto = (testo, nome) => {
      const { intestazioni, righe } = leggiTabella(testo);
      if (!intestazioni.length) { setImp((x) => ({ ...x, esito: { errore: "Non ho trovato colonne leggibili." } })); return; }
      setImp((x) => ({ ...x, intestazioni, righe, nomeFile: nome || "incollato", mappa: indovinaMappa(x.tipo, intestazioni), esito: null }));
    };

    const daFile = (file) => {
      const r = new FileReader();
      if (eExcel(file.name)) {
        r.onload = () => {
          try {
            const fogli = fogliDaExcel(r.result);
            if (!fogli.length) { setImp((x) => ({ ...x, esito: { errore: "Il foglio di calcolo non contiene righe leggibili." } })); return; }
            const scelto = fogli.slice().sort((a, b) => b.righe - a.righe)[0];
            const { intestazioni, righe } = leggiTabella(scelto.csv);
            setImp((x) => ({
              ...x, intestazioni, righe, fogli, foglio: scelto.nome,
              nomeFile: `${file.name} · foglio "${scelto.nome}"`,
              mappa: indovinaMappa(x.tipo, intestazioni), esito: null,
            }));
          } catch (e) {
            setImp((x) => ({ ...x, esito: { errore: "Non sono riuscito ad aprire il foglio di calcolo: prova a salvarlo in CSV." } }));
          }
        };
        r.onerror = () => setImp((x) => ({ ...x, esito: { errore: "Non sono riuscito a leggere il file." } }));
        r.readAsArrayBuffer(file);
        return;
      }
      r.onload = () => { setImp((x) => ({ ...x, fogli: [], foglio: "" })); caricaTesto(r.result, file.name); };
      r.onerror = () => setImp((x) => ({ ...x, esito: { errore: "Non sono riuscito a leggere il file." } }));
      r.readAsText(file);
    };

    const cambiaFoglio = (nome) => {
      const f = (imp.fogli || []).find((x) => x.nome === nome);
      if (!f) return;
      const { intestazioni, righe } = leggiTabella(f.csv);
      setImp((x) => ({ ...x, intestazioni, righe, foglio: nome, nomeFile: `${x.nomeFile.split(" · ")[0]} · foglio "${nome}"`, mappa: indovinaMappa(x.tipo, intestazioni), esito: null }));
    };

    /* Un collegamento normale a Google Fogli restituisce una pagina, non i dati:
       lo converto nella forma che esporta il CSV del foglio aperto. */
    const urlCsv = (u) => {
      const g = /docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/.exec(u || "");
      if (!g) return u;
      if (/output=csv|format=csv/.test(u)) return u;
      const gid = (/[#&?]gid=(\d+)/.exec(u) || [])[1] || "0";
      return `https://docs.google.com/spreadsheets/d/${g[1]}/export?format=csv&gid=${gid}`;
    };

    const daUrl = async () => {
      if (!imp.url) return;
      const u = urlCsv(imp.url.trim());
      setImp((x) => ({ ...x, url: u, esito: { attesa: true } }));
      try {
        const r = await fetch(u);
        if (!r.ok) throw new Error(String(r.status));
        const t = await r.text();
        if (/^\s*</.test(t)) throw new Error("html");
        caricaTesto(t, u);
      } catch (e) {
        const dentroClaude = /claude|anthropic/i.test(location.hostname || "");
        setImp((x) => ({
          ...x,
          esito: {
            errore: dentroClaude
              ? "Qui dentro il collegamento diretto non può funzionare: l'app gira in una finestra isolata che può contattare solo pochi indirizzi, e Google non è tra questi. Scarica il foglio (File → Scarica → CSV) e caricalo con il campo qui sopra: il risultato è identico."
              : "La sorgente non risponde o non consente la lettura da un'altra pagina. Se è un foglio Google, verifica che sia pubblicato sul web in formato CSV; altrimenti scaricalo e caricalo come file.",
          },
        }));
      }
    };

    const valore = (riga, k) => {
      const col = imp.mappa[k];
      if (!col) return "";
      const i = imp.intestazioni.indexOf(col);
      return i < 0 ? "" : riga[i];
    };
    const campoNum = (k) => (cfg.campi.find((c) => c.k === k) || {}).num;
    /* Il separatore decimale lo decido guardando tutta la colonna, non la singola cella:
       "1.234" da solo è ambiguo, ma se nella stessa colonna c'è "12,5" il file è all'italiana. */
    const cacheCol = {};
    const decimaleCol = (k) => {
      if (cacheCol["n" + k] !== undefined) return cacheCol["n" + k];
      let virgola = 0, punto = 0;
      imp.righe.forEach((r) => {
        const t = String(valore(r, k) || "").replace(/[^\d,.\-]/g, "");
        if (/\d,\d{1,2}$|\d,\d{4,}$|\.\d{3},\d/.test(t)) virgola++;
        if (/\d\.\d{1,2}$|\d\.\d{4,}$|,\d{3}\.\d/.test(t)) punto++;
      });
      return (cacheCol["n" + k] = virgola > punto ? "," : punto > virgola ? "." : "");
    };
    const numCol = (k, s) => {
      if (typeof s === "number") return isFinite(s) ? s : 0;
      const dec = decimaleCol(k);
      if (!dec) return parseNumIntl(s);
      let t = String(s || "").replace(/[^\d,.\-]/g, "");
      t = dec === "," ? t.replace(/\./g, "").replace(",", ".") : t.replace(/,/g, "");
      const n = parseFloat(t);
      return isFinite(n) ? n : 0;
    };
    /* Stessa logica per le date: se in colonna c'è un 25/03 è giorno/mese, se c'è un 03/25 è mese/giorno */
    const ordineCol = (k) => {
      if (cacheCol["d" + k] !== undefined) return cacheCol["d" + k];
      let gm = 0, mg = 0;
      imp.righe.forEach((r) => {
        const m2 = String(valore(r, k) || "").trim().match(/^(\d{1,2})[-/.](\d{1,2})[-/.]\d{4}/);
        if (!m2) return;
        if (+m2[1] > 12) gm++;
        if (+m2[2] > 12) mg++;
      });
      return (cacheCol["d" + k] = mg > gm ? "md" : gm > mg ? "gm" : "");
    };
    const dataCol = (riga, k) => normalizzaData(valore(riga, k), ordineCol(k));
    const leggi = (riga, k) => (campoNum(k) ? numCol(k, valore(riga, k)) : String(valore(riga, k) || "").trim());

    const anteprima = imp.righe.slice(0, 4).map((r) => cfg.campi.map((c) => leggi(r, c.k)));

    const importa = () => {
      const cid = imp.clienteId;
      let creati = 0, aggiornati = 0, saltati = 0, dettaglio = "";
      const nuoviCli = [];
      const norm = (s) => String(s || "").trim().toLowerCase();
      const ambigui = new Set();
      /* Senza cliente scelto aggancio una campagna solo se il nome è unico in tutto l'archivio:
         due clienti possono avere una campagna con lo stesso nome. */
      const indiceCamp = (lista, nome) => {
        const idx = lista.map((k, j) => (norm(k.nome) === norm(nome) && (!cid || k.clienteId === cid) ? j : -1)).filter((j) => j > -1);
        if (idx.length > 1 && !cid) { ambigui.add(nome); return -2; }
        return idx.length ? idx[0] : -1;
      };
      const trovaCamp = (lista, nome) => { const j = indiceCamp(lista, nome); return j >= 0 ? lista[j] : null; };

      const ndFinale = ((d) => {
        const nd = { ...d };

        if (imp.tipo === "campagne") {
          const lista = [...nd.campagne];
          imp.righe.forEach((r) => {
            const nome = leggi(r, "nome");
            if (!nome) { saltati++; return; }
            const i = indiceCamp(lista, nome);
            if (i === -2) { saltati++; return; }
            const metriche = {};
            ["spesa", "impression", "click", "clickLink", "lead", "conversioni", "ricavi",
              "valoreContratto", "incassato", "frequenza", "budget"].forEach((k) => {
              if (imp.mappa[k]) metriche[k] = leggi(r, k);
            });
            if (imp.mappa.linea && leggi(r, "linea")) metriche.linea = leggi(r, "linea");
            /* Date, stato e impostazioni: senza queste le campagne nascono tutte oggi e tutte attive */
            ["inizio", "fine"].forEach((k) => {
              if (imp.mappa[k] && dataCol(r, k)) metriche[k] = dataCol(r, k);
            });
            if (imp.mappa.budgetGiorno) metriche.budgetGiorno = leggi(r, "budgetGiorno");
            if (imp.mappa.piattaforma) {
              const pl = norm(leggi(r, "piattaforma"));
              const tr = PIATTAFORME.find((x) => x.v === pl || norm(x.l) === pl);
              if (tr) metriche.piattaforma = tr.v;
            }
            if (imp.mappa.obiettivo) {
              const ob = norm(leggi(r, "obiettivo"));
              const tr = OBIETTIVI.find((x) => x.v === ob || norm(x.l) === ob);
              if (tr) metriche.obiettivo = tr.v;
            }
            if (imp.mappa.stato) {
              const st = norm(leggi(r, "stato"));
              const diretto = STATI_CAMPAGNA.find((x) => x.v === st || norm(x.l) === st);
              /* Riconosco anche gli stati come li scrive Meta */
              const meta = /active|attiv/.test(st) ? "attiva"
                : /paused|pausa|spenta/.test(st) ? "pausa"
                  : /complet|conclus|ended|archiv/.test(st) ? "conclusa"
                    : /error|unsettled|pagament/.test(st) ? "errore" : "";
              const scelto = diretto ? diretto.v : meta;
              if (scelto) metriche.stato = scelto;
            }
            if (i > -1) { lista[i] = { ...lista[i], ...metriche }; aggiornati++; }
            else if (cid) {
              lista.unshift({
                id: uid(), clienteId: cid, nome, piattaforma: "altro", obiettivo: "vendite", stato: "attiva",
                inizio: oggi(), fine: piu(30), linea: "", budgetGiorno: 0, budget: 0, spesa: 0, impression: 0, click: 0, clickLink: 0, lead: 0, conversioni: 0, ricavi: 0, valoreContratto: 0, incassato: 0, frequenza: 0,
                bmk: { ...BMK_DEFAULT.vendite }, ...metriche,
              });
              creati++;
            } else saltati++;
          });
          nd.campagne = lista;
        }

        if (imp.tipo === "creativi") {
          const lista = [...nd.creativi];
          imp.righe.forEach((r) => {
            const nome = leggi(r, "nome");
            if (!nome) { saltati++; return; }
            const nomeCamp = leggi(r, "campagna");
            const camp = trovaCamp(nd.campagne, nomeCamp);
            const metriche = {};
            ["spesa", "impression", "click", "clickLink", "lead", "conversioni", "ricavi", "valoreContratto", "incassato", "frequenza"].forEach((k) => {
              if (imp.mappa[k]) metriche[k] = leggi(r, k);
            });
            /* Il livello lo deduco dalla colonna, o dal fatto che sia indicato un ad set padre */
            const nomeAdset = imp.mappa.adset ? leggi(r, "adset") : "";
            const tipoLetto = norm(imp.mappa.tipo ? leggi(r, "tipo") : "");
            const tipo = /annunc|ad$|ads$|creativ/.test(tipoLetto) ? "annuncio"
              : /adset|gruppo/.test(tipoLetto) ? "adset"
                : (nomeAdset && norm(nomeAdset) !== norm(nome)) ? "annuncio" : "adset";
            const padre = nomeAdset
              ? lista.find((c) => c.tipo === "adset" && norm(c.nome) === norm(nomeAdset) && (!camp || c.campagnaId === camp.id))
              : null;
            if (padre) metriche.adsetId = padre.id;
            metriche.tipo = tipo;
            const i = lista.findIndex((c) => norm(c.nome) === norm(nome) && (!camp || c.campagnaId === camp.id));
            if (i > -1) { lista[i] = { ...lista[i], ...metriche }; aggiornati++; }
            else if (camp) {
              lista.unshift({
                id: uid(), clienteId: camp.clienteId, campagnaId: camp.id, adsetId: "", tipo: "adset", nome, stato: "attiva",
                spesa: 0, impression: 0, click: 0, conversioni: 0, ricavi: 0, valoreContratto: 0, incassato: 0, frequenza: 0, note: "", ...metriche,
              });
              creati++;
            } else saltati++;
          });
          nd.creativi = lista;
        }

        if (imp.tipo === "clienti") {
          const lista = [...nd.clienti];
          const statoValido = (x) => (STATI_CLIENTE.some((st) => st.v === norm(x)) ? norm(x)
            : STATI_CLIENTE.find((st) => norm(st.l) === norm(x)) ? STATI_CLIENTE.find((st) => norm(st.l) === norm(x)).v
              : "onboarding");
          imp.righe.forEach((r) => {
            const nome = leggi(r, "nome");
            if (!nome) { saltati++; return; }
            const campi = {};
            ["settore", "citta", "referenteAmm", "emailAmm", "ragioneSociale", "piva", "accountAdv", "cognome", "note"].forEach((k) => {
              if (imp.mappa[k] && leggi(r, k)) campi[k] = leggi(r, k);
            });
            ["fee", "budgetMese", "aov", "margineLordo"].forEach((k) => {
              if (imp.mappa[k]) campi[k] = leggi(r, k);
            });
            ["creato", "dataContratto"].forEach((k) => {
              if (imp.mappa[k] && dataCol(r, k)) campi[k] = dataCol(r, k);
            });
            if (imp.mappa.stato && leggi(r, "stato")) campi.stato = statoValido(leggi(r, "stato"));
            /* Stesso nome sotto lo stesso padre: è lo stesso cliente, lo aggiorno */
            const i = lista.findIndex((z) => norm(z.nome) === norm(nome) && (z.padreId || "") === (cid || ""));
            if (i > -1) { lista[i] = { ...lista[i], ...campi }; aggiornati++; }
            else {
              const z = {
                id: uid(), nome, padreId: cid || "", archiviato: "no", dataContratto: "",
                settore: "", citta: "", stato: "onboarding", fee: 0, budgetMese: 0,
                margineLordo: 0, resi: 0, aov: 0, aoa: {}, logo: "", note: "", creato: oggi(),
                ragioneSociale: "", piva: "", codiceFiscale: "", indirizzo: "", cap: "", cittaFatt: "",
                provincia: "", paese: "Italia", sdi: "", pec: "", emailAmm: "", referenteAmm: "",
                termini: "30", metodo: "bonifico", ...campi,
              };
              lista.unshift(z);
              nuoviCli.push({ id: z.id, nome: z.nome });
              creati++;
            }
          });
          nd.clienti = lista;
        }

        if (imp.tipo === "fasi") {
          const lista = [...(nd.periodi || [])];
          const mancanti = new Set();
          imp.righe.forEach((r) => {
            const nomeCamp = leggi(r, "campagna");
            const dal = dataCol(r, "inizio");
            if (!nomeCamp || !dal) { saltati++; return; }
            const camp = trovaCamp(nd.campagne, nomeCamp);
            if (!camp) { saltati++; mancanti.add(nomeCamp); return; }
            const tipoLetto = norm(leggi(r, "tipo"));
            const tipo = ["test", "matura", "attiva"].includes(tipoLetto) ? tipoLetto : "attiva";
            const metriche = {};
            ["spesa", "impression", "click", "clickLink", "lead", "conversioni", "ricavi", "frequenza"].forEach((k) => {
              if (imp.mappa[k]) metriche[k] = leggi(r, k);
            });
            const fine = dataCol(r, "fine") || dal;
            /* Stessa campagna e stessa data di inizio: è la stessa finestra, la aggiorno */
            const i = lista.findIndex((pz) => pz.campagnaId === camp.id && pz.inizio === dal);
            if (i > -1) { lista[i] = { ...lista[i], tipo, fine, ...metriche }; aggiornati++; }
            else {
              lista.push({
                id: uid(), clienteId: camp.clienteId, campagnaId: camp.id, tipo, inizio: dal, fine,
                spesa: 0, impression: 0, click: 0, clickLink: 0, lead: 0, conversioni: 0, ricavi: 0,
                valoreContratto: 0, incassato: 0, frequenza: 0, note: "", ...metriche,
              });
              creati++;
            }
          });
          nd.periodi = lista;
          if (mancanti.size) dettaglio = `Campagne non trovate su questo cliente: ${[...mancanti].slice(0, 6).join(" · ")}${mancanti.size > 6 ? ` e altre ${mancanti.size - 6}` : ""}. Importa prima il file delle campagne, con tipo "Campagne pubblicitarie".`;
        }

        if (imp.tipo === "vendite") {
          if (!cid) { saltati = imp.righe.length; }
          else {
            const lista = [...nd.vendite];
            imp.righe.forEach((r) => {
              const periodo = normalizzaMese(valore(r, "periodo"));
              if (!periodo) { saltati++; return; }
              const campi = {};
              ["lead", "chiamate", "discoveryFissate", "discoverySvolte", "demoFissate", "demoSvolte", "offerte", "spesa", "slot", "oreContatto"]
                .forEach((k) => { if (imp.mappa[k]) campi[k] = leggi(r, k); });
              const i = lista.findIndex((v) => v.clienteId === cid && v.periodo === periodo);
              if (i > -1) { lista[i] = { ...lista[i], ...campi }; aggiornati++; }
              else {
                lista.unshift({
                  id: uid(), clienteId: cid, periodo, lead: 0, chiamate: 0, discoveryFissate: 0, discoverySvolte: 0,
                  demoFissate: 0, demoSvolte: 0, offerte: 0, chiusure: 0, valoreContratto: 0, incassato: 0, spesa: 0,
                  slot: 0, oreContatto: 0, ...campi,
                });
                creati++;
              }
            });
            nd.vendite = lista;
          }
        }

        if (imp.tipo === "chiusure") {
          if (!cid) { saltati = imp.righe.length; }
          else {
            const lista = [...nd.chiusure];
            imp.righe.forEach((r) => {
              const nome = leggi(r, "nome");
              const dataC = dataCol(r, "data") || oggi();
              if (!nome) { saltati++; return; }
              if (lista.some((c) => c.clienteId === cid && norm(c.nome) === norm(nome) && c.data === dataC)) { saltati++; return; }
              const nomeCamp = leggi(r, "campagna");
              const camp = nd.campagne.find((k) => norm(k.nome) === norm(nomeCamp) && k.clienteId === cid);
              const fun = String(leggi(r, "funnel") || "").toLowerCase();
              lista.unshift({
                id: uid(), clienteId: cid, data: dataC, nome,
                azienda: leggi(r, "azienda"), email: leggi(r, "email"), telefono: leggi(r, "telefono"),
                funnel: (FUNNEL_PROV.find((f) => fun.includes(f.v) || fun.includes(f.l.toLowerCase())) || { v: "altro" }).v,
                campagnaId: camp ? camp.id : "", adset: leggi(r, "adset"), annuncio: leggi(r, "annuncio"),
                closer: leggi(r, "closer"), valoreContratto: leggi(r, "valoreContratto"), incassato: leggi(r, "incassato"), note: "",
              });
              creati++;
            });
            nd.chiusure = lista;
          }
        }

        return nd;
      })(data);
      setData(ndFinale);

      if (ambigui.size) {
        dettaglio = `${dettaglio ? dettaglio + " " : ""}Nomi presenti su più clienti, non agganciati: ${[...ambigui].slice(0, 6).join(" · ")}${ambigui.size > 6 ? ` e altri ${ambigui.size - 6}` : ""}. Scegli il cliente in alto e ripeti l'importazione.`;
      }
      setImp((x) => ({ ...x, esito: { creati, aggiornati, saltati, nuoviCli, padre: cid, dettaglio } }));
    };

    const salvaProfilo = () => {
      chiedi({
        titolo: "Salva il profilo di importazione",
        testo: "Memorizza tipo, cliente e abbinamento delle colonne: la prossima volta ti basta sceglierlo e caricare il file.",
        etichetta: "Salva",
        campo: { label: "Nome del profilo", valore: "" },
        azione: (nome) => upd("profili", (l) => [{ id: uid(), nome, tipo: imp.tipo, clienteId: imp.clienteId, url: imp.url, mappa: imp.mappa, ultimo: oggi() }, ...l]),
      });
    };
    const caricaProfilo = (id) => {
      const p = data.profili.find((x) => x.id === id);
      if (!p) { setImp((x) => ({ ...x, profilo: "" })); return; }
      setImp((x) => ({ ...x, profilo: id, tipo: p.tipo, clienteId: p.clienteId, url: p.url, mappa: x.intestazioni.length ? p.mappa : x.mappa }));
    };

    return (
      <>
        <section className="card">
          <header>
            <div className="ctrl">
              <select value={imp.tipo} onChange={(e) => setImp({ ...imp, tipo: e.target.value, mappa: imp.intestazioni.length ? indovinaMappa(e.target.value, imp.intestazioni) : {}, esito: null })} aria-label="Cosa importi">
                {Object.entries(CAMPI_IMPORT).map(([k, v]) => <option key={k} value={k}>{v.l}</option>)}
              </select>
              <select value={imp.clienteId} onChange={(e) => setImp({ ...imp, clienteId: e.target.value })}
                aria-label={imp.tipo === "clienti" ? "Cliente madre" : "Cliente di destinazione"}>
                <option value="">{imp.tipo === "clienti" ? "— come clienti principali —" : "— scegli il cliente —"}</option>
                {(imp.tipo === "clienti" ? attivi.filter((z) => !z.padreId) : attivi)
                  .map((z) => <option key={z.id} value={z.id}>{imp.tipo === "clienti" ? `sotto ${z.nome}` : z.nome}</option>)}
              </select>
              {data.profili.length > 0 && (
                <select value={imp.profilo} onChange={(e) => caricaProfilo(e.target.value)} aria-label="Profilo salvato">
                  <option value="">Profilo salvato…</option>
                  {data.profili.map((p) => <option key={p.id} value={p.id}>{p.nome}</option>)}
                </select>
              )}
            </div>
          </header>
          <div style={{ padding: "16px 18px", display: "grid", gap: 12 }}>
            <div className="f">
              <label htmlFor="imp-file">Carica un file dal tuo computer — Excel, CSV o TSV</label>
              <div className="drop" data-sopra={trascina ? "1" : "0"}
                onDragOver={(e) => { e.preventDefault(); setTrascina(true); }}
                onDragLeave={() => setTrascina(false)}
                onDrop={(e) => {
                  e.preventDefault(); setTrascina(false);
                  const fl = e.dataTransfer.files && e.dataTransfer.files[0];
                  if (fl) daFile(fl);
                }}>
                <input id="imp-file" type="file" accept=".csv,.tsv,.txt,text/csv,.xlsx,.xlsm,.xls,.ods"
                  onChange={(e) => { const fl = e.target.files && e.target.files[0]; if (fl) daFile(fl); e.target.value = ""; }} />
                <span className="sub">…oppure trascinalo qui dentro</span>
              </div>
              <span className="hint">
                {{
                  clienti: "Da Google Fogli: File → Scarica → CSV. Va bene anche un .xlsx esportato da Excel o Numbers.",
                  fasi: "Il file delle finestre prodotto dallo script di estrazione, o un foglio con una riga per campagna e settimana.",
                  campagne: "Il report campagne esportato da Meta, Google o dallo script di estrazione.",
                  creativi: "Il report per ad set o annuncio esportato dalla piattaforma.",
                  vendite: "Un foglio con una riga per mese di attività commerciale.",
                  chiusure: "Un foglio con una riga per cliente chiuso.",
                }[imp.tipo] || "Il file esportato dalla piattaforma o dal tuo foglio."}
              </span>
            </div>
            <div className="f">
              <label htmlFor="imp-testo">Oppure incolla qui le righe copiate da un foglio</label>
              <textarea id="imp-testo" rows={4} value={imp.testo} placeholder="Nome campagna	Speso	Impression	Click…"
                onChange={(e) => setImp({ ...imp, testo: e.target.value })} />
              <button className="btn ghost mini" style={{ alignSelf: "flex-start", marginTop: 8 }}
                onClick={() => caricaTesto(imp.testo, "incollato")}>Leggi il testo incollato</button>
            </div>
            <div className="f">
              <label htmlFor="imp-url">Sorgente remota (foglio Google pubblicato in CSV, o automazione che restituisce CSV)</label>
              <div className="riga">
                <input id="imp-url" value={imp.url} placeholder="https://…" onChange={(e) => setImp({ ...imp, url: e.target.value })} />
                <button className="btn ghost" onClick={daUrl}>Sincronizza ora</button>
              </div>
              <span className="hint">
                Dentro Claude il collegamento diretto non funziona: la finestra dell'app può contattare solo pochi indirizzi, e Google
                non è tra questi. Funziona invece il file scaricato. Fuori da Claude funziona se la sorgente consente la lettura da browser;
                incollando il link normale del foglio lo converto comunque nella forma che esporta il CSV.
              </span>
            </div>
          </div>
        </section>

        {imp.esito && imp.esito.errore && (
          <section className="card"><div className="empty"><b>Non ha funzionato</b>{imp.esito.errore}</div></section>
        )}

        {pronti && (
          <>
            <section className="card">
              <header>
                <h2>Abbinamento colonne</h2>
                <span className="sp eyebrow">{imp.righe.length} righe da {imp.nomeFile}</span>
                {(imp.fogli || []).length > 1 && (
                  <select value={imp.foglio} onChange={(e) => cambiaFoglio(e.target.value)} aria-label="Foglio del file"
                    style={{ width: "auto", minWidth: 150, fontSize: 12.5, padding: "6px 8px" }}>
                    {imp.fogli.map((f) => <option key={f.nome} value={f.nome}>{f.nome} ({f.righe} righe)</option>)}
                  </select>
                )}
                <button className="btn ghost mini" onClick={salvaProfilo}>Salva profilo</button>
              </header>
              <div className="fields">
                {cfg.campi.map((c) => (
                  <div className="f" key={c.k}>
                    <label htmlFor={"m-" + c.k}>{c.l}</label>
                    <select id={"m-" + c.k} value={imp.mappa[c.k] || ""}
                      onChange={(e) => setImp({ ...imp, mappa: { ...imp.mappa, [c.k]: e.target.value } })}>
                      <option value="">— non importare —</option>
                      {imp.intestazioni.map((h, i) => <option key={h + i} value={h}>{h}</option>)}
                    </select>
                  </div>
                ))}
              </div>
            </section>

            <section className="card">
              <header><h2>Anteprima</h2><span className="sp eyebrow">prime righe, come verranno lette</span></header>
              <div className="wrap">
                <table className="tbl">
                  <thead><tr>{cfg.campi.map((c) => <th key={c.k} className={c.num ? "r" : ""}>{c.l}</th>)}</tr></thead>
                  <tbody>
                    {anteprima.map((r, i) => (
                      <tr key={i}>{r.map((v, j) => (
                        <td key={j} className={cfg.campi[j].num ? "r" : ""}>
                          {cfg.campi[j].num ? (v ? fmt(v, 2) : "—") : (v || <span className="sub">—</span>)}
                        </td>
                      ))}</tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{ padding: "14px 16px", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                <button className="btn" onClick={importa}>Importa {imp.righe.length} righe</button>
                <button className="btn ghost" onClick={() => setImp({ ...imp, intestazioni: [], righe: [], esito: null })}>Annulla</button>
                <span className="sub">
                  {imp.tipo === "clienti" ? "Scegli sopra il cliente madre e ogni riga diventa un suo sottocliente, con scheda, campagne, funnel e ciclo di vendita propri. Lasciando “come clienti principali” vengono creati al primo livello. Chi ha già lo stesso nome sotto lo stesso madre viene aggiornato, non duplicato."
                    : imp.tipo === "fasi" ? "Le finestre si agganciano alla campagna per nome esatto. Stessa campagna e stessa data di inizio significa stessa finestra: viene aggiornata, non duplicata."
                    : imp.tipo === "chiusure" ? "Le righe già presenti (stesso nome e data) vengono saltate."
                    : "Le righe già esistenti vengono aggiornate, le altre create."}
                </span>
              </div>
            </section>
          </>
        )}

        {imp.esito && !imp.esito.errore && !imp.esito.attesa && (
          <article className="al" data-g={imp.esito.saltati ? "attenzione" : "attenzione"}>
            <div className="meta"><span className="tag" data-k="attivo">fatto</span></div>
            <h3>{int(imp.esito.aggiornati)} aggiornati · {int(imp.esito.creati)} creati · {int(imp.esito.saltati)} saltati</h3>
            {imp.esito.nuoviCli && imp.esito.nuoviCli.length > 0 && (
              <dl>
                <dt>Schede create</dt>
                <dd>
                  {imp.esito.padre
                    ? `Ognuno è ora un sottocliente di ${nomeCli(imp.esito.padre)}, con campagne, funnel, messaggio, ciclo di vendita, chiusure e task propri.`
                    : "Ognuno è ora un cliente con la sua struttura completa."}
                  <div className="riga" style={{ gap: 6, flexWrap: "wrap", marginTop: 9 }}>
                    {imp.esito.nuoviCli.slice(0, 24).map((z) => (
                      <button key={z.id} className="btn ghost mini" onClick={() => cambiaAccount(z.id)}>{z.nome}</button>
                    ))}
                    {imp.esito.nuoviCli.length > 24 && <span className="sub">e altri {imp.esito.nuoviCli.length - 24}</span>}
                  </div>
                </dd>
              </dl>
            )}
            {imp.esito.dettaglio && (
              <dl>
                <dt>Cosa manca</dt>
                <dd>{imp.esito.dettaglio}</dd>
              </dl>
            )}
            {imp.esito.saltati > 0 && (
              <dl>
                <dt>Perché saltati</dt>
                <dd>Righe senza nome, oppure senza un cliente di destinazione scelto, oppure — per ad set e fasi — con una campagna che non esiste ancora qui dentro. Importa prima le campagne, poi il resto. Per i clienti, l’unica causa di scarto è la riga senza nome: controlla che la colonna del nome sia abbinata nella tabella qui sopra.</dd>
              </dl>
            )}
          </article>
        )}

        <section className="card">
          <header><h2>Come renderlo quasi automatico</h2></header>
          <div style={{ padding: "14px 16px", fontSize: 13.3, color: "var(--ink2)" }}>
            Salva un profilo per ogni esportazione ricorrente: la volta dopo scegli il profilo, trascini il file e importi, senza rifare l'abbinamento.
            Le piattaforme permettono anche invii programmati via email dei report: con Make o Zapier puoi far scrivere quell'allegato in un foglio,
            pubblicare il foglio in CSV e incollarne qui l'indirizzo, così ti basta premere Sincronizza.
            <br /><br />
            Quello che non si può fare da qui è collegarsi da soli alle API di Meta o Google: servono chiavi di accesso che vanno custodite
            su un server, e un processo che giri anche quando questa pagina è chiusa. Finché il software vive nel browser, l'ultimo passo resta un tuo click.
          </div>
        </section>
      </>
    );
  };

  /* ---------- backup ---------- */
  const SezioneBackup = () => {
    const conta = {
      Clienti: data.clienti.length, Contatti: data.contatti.length, Campagne: data.campagne.length,
      "Ad set e annunci": data.creativi.length, Trattative: data.trattative.length, Chiusure: data.chiusure.length,
      "Periodi commerciali": data.vendite.length, Funnel: data.funnel.length, Task: data.task.length,
      Annotazioni: data.annotazioni.length, Esperimenti: data.esperimenti.length, Automazioni: data.automazioni.length,
    };
    const peso = JSON.stringify(data).length;
    const giorniDaFile = data.backup.ultimoFile
      ? Math.floor((new Date(oggi()) - new Date(data.backup.ultimoFile.slice(0, 10))) / 864e5) : null;
    const daScaricare = giorniDaFile === null || giorniDaFile >= 7;

    const esporta = esportaBackup;

    const leggiFile = (file) => {
      const r = new FileReader();
      r.onload = () => {
        try {
          const j = JSON.parse(r.result);
          const dd = j.dati || j;
          if (!dd || !Array.isArray(dd.clienti) || !Array.isArray(dd.campagne)) throw new Error("struttura");
          setDaRipristinare({
            nome: file.name, generato: j.generato || "", dati: dd,
            riepilogo: { clienti: dd.clienti.length, campagne: dd.campagne.length, task: (dd.task || []).length, chiusure: (dd.chiusure || []).length },
          });
        } catch (e) {
          setDaRipristinare({ errore: "Questo file non è un backup leggibile: deve essere il .json scaricato da qui." });
        }
      };
      r.onerror = () => setDaRipristinare({ errore: "Non sono riuscito a leggere il file." });
      r.readAsText(file);
    };

    const applica = () => {
      sbloccaArchivio();
      setData((attuale) => ({ ...normalizza(daRipristinare.dati), aiKey: (attuale && attuale.aiKey) || daRipristinare.dati.aiKey || "" }));
      setDaRipristinare(null); setCliAtt(""); setVista("cruscotto");
    };

    const ripristinaCopia = async (s) => {
      let dd = null;
      try {
        const r = await window.storage.get(s.chiave);
        dd = JSON.parse(r.value);
      } catch (e) {
        chiedi({ titolo: "Copia non leggibile", testo: "Questa copia non è più recuperabile: usa un backup scaricato.", etichetta: "Ho capito", azione: () => {} });
        return;
      }
      chiedi({
        titolo: `Ripristinare la copia del ${dataIt(s.chiave.slice(-10))}?`,
        testo: "I dati attuali vengono sostituiti con quelli di quel giorno. L'operazione non si può annullare.",
        etichetta: "Ripristina", pericolo: true,
        azione: () => { sbloccaArchivio(); setData((attuale) => ({ ...normalizza(dd), aiKey: (attuale && attuale.aiKey) || dd.aiKey || "" })); setCliAtt(""); setVista("cruscotto"); },
      });
    };

    return (
      <>
        {daScaricare && (
          <article className="al" data-g={giorniDaFile === null ? "critica" : "attenzione"}>
            <div className="meta"><span className="tag" data-k="late">da fare</span>
              <span className="misura">{giorniDaFile === null ? "nessun file scaricato" : `ultimo file ${giorniDaFile} giorni fa`}</span></div>
            <h3>Scarica una copia sul tuo computer</h3>
            <dl>
              <dt>Cosa succede</dt><dd>Tutti i dati vivono nell'archivio di questo browser. Se cancelli i dati di navigazione, cambi computer o il browser fa pulizia, spariscono — e le copie automatiche spariscono con loro, perché stanno nello stesso posto.</dd>
              <dt>Cosa fare</dt><dd>Un file al mese è il minimo, uno a settimana è la cosa giusta. Sono pochi secondi e ti mette al riparo dal caso peggiore.</dd>
            </dl>
          </article>
        )}

        <div className="two">
          <section className="card">
            <header><h2>Salvataggio manuale</h2><span className="sp" /><button className="btn" onClick={esporta}>Scarica backup</button></header>
            <table className="tbl"><tbody>
              <tr><td className="sub">Ultimo file scaricato</td><td>{data.backup.ultimoFile ? new Date(data.backup.ultimoFile).toLocaleString(LOCALE) : <span className="neg">mai</span>}</td></tr>
              <tr><td className="sub">Peso stimato</td><td className="mono">{fmt(peso / 1024, 0)} kB</td></tr>
              <tr><td className="sub">Formato</td><td>file .json, leggibile e riutilizzabile</td></tr>
            </tbody></table>
            <div className="nota">Il file contiene tutto: clienti, campagne, storico, funnel, task, automazioni e i loghi caricati.</div>
          </section>

          <section className="card">
            <header><h2>Copia automatica</h2><span className="sp" />
              <button className="btn ghost mini" data-on={data.backup.auto === "si" ? "1" : "0"}
                onClick={() => setData((d) => ({ ...d, backup: { ...d.backup, auto: d.backup.auto === "si" ? "no" : "si" } }))}>
                {data.backup.auto === "si" ? "Attiva" : "Disattivata"}
              </button>
            </header>
            <table className="tbl"><tbody>
              <tr><td className="sub">Frequenza</td><td>una volta al giorno, al primo utilizzo</td></tr>
              <tr><td className="sub">Ultima copia</td><td>{data.backup.ultimoAuto ? dataIt(data.backup.ultimoAuto) : "—"}</td></tr>
              <tr><td className="sub">Copie conservate</td><td>{data.snapshots.length} di 6</td></tr>
            </tbody></table>
            {data.backup.errore && <div className="nota neg">{data.backup.errore}</div>}
            <div className="nota">Le copie automatiche restano dentro il browser: ti salvano da un errore tuo o da dati rovinati, non dalla perdita del dispositivo.</div>
          </section>
        </div>

        <section className="card">
          <header><h2>Copie disponibili</h2><span className="sp eyebrow">le più recenti in alto</span></header>
          <table className="tbl">
            <thead><tr><th>Copia</th><th className="r">Peso</th><th className="r">Azioni</th></tr></thead>
            <tbody>
              {data.snapshots.map((s) => (
                <tr key={s.chiave}>
                  <td><b>{dataIt(s.chiave.slice(-10))}</b><div className="sub">{new Date(s.quando).toLocaleString(LOCALE)}</div></td>
                  <td className="r">{fmt(s.peso / 1024, 0)} kB</td>
                  <td className="r">
                    <button className="btn ghost mini" onClick={() => ripristinaCopia(s)}>Ripristina</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {!data.snapshots.length && <div className="empty"><b>Nessuna copia ancora</b>La prima viene creata automaticamente oggi stesso, se la copia automatica è attiva.</div>}
        </section>

        <section className="card">
          <header><h2>Ripristino da file</h2></header>
          <div style={{ padding: "16px 18px", display: "flex", flexDirection: "column", gap: 12, maxWidth: 520 }}>
            <input type="file" accept="application/json,.json"
              onChange={(e) => { const fl = e.target.files && e.target.files[0]; if (fl) leggiFile(fl); e.target.value = ""; }} />
            {daRipristinare && daRipristinare.errore && <div className="neg">{daRipristinare.errore}</div>}
            {daRipristinare && daRipristinare.dati && (
              <div className="al" data-g="attenzione">
                <h3>Confermi il ripristino?</h3>
                <dl>
                  <dt>File</dt><dd>{daRipristinare.nome}{daRipristinare.generato ? ` · generato il ${new Date(daRipristinare.generato).toLocaleString(LOCALE)}` : ""}</dd>
                  <dt>Contiene</dt><dd>{daRipristinare.riepilogo.clienti} clienti, {daRipristinare.riepilogo.campagne} campagne, {daRipristinare.riepilogo.chiusure} chiusure, {daRipristinare.riepilogo.task} task.</dd>
                  <dt>Attenzione</dt><dd>I dati attuali vengono sostituiti e non tornano indietro. Se non sei sicuro, scarica prima un backup di adesso.</dd>
                </dl>
                <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                  <button className="btn ghost" onClick={() => setDaRipristinare(null)}>Annulla</button>
                  <button className="btn" onClick={applica}>Ripristina i dati</button>
                </div>
              </div>
            )}
          </div>
        </section>

        <section className="card">
          <header><h2>Cosa contiene l'archivio</h2></header>
          <div className="wrap">
            <table className="tbl"><tbody>
              {Object.entries(conta).map(([k, v]) => (
                <tr key={k}><td>{k}</td><td className="r">{int(v)}</td></tr>
              ))}
            </tbody></table>
          </div>
        </section>
      </>
    );
  };

  /* ---------- funnel builder ---------- */
  const SezioneFunnel = () => {
    const lista = attivo ? S.funnel : data.funnel;

    /* I modelli si gestiscono anche quando non esiste ancora nessun funnel:
       è il momento in cui servono di più. */
    const schedaModelli = (corrente) => (
      <section className="card">
        <header>
          <h2>Modelli di percorso</h2>
          <span className="sp eyebrow">
          {(data.modelli || []).length ? `${data.modelli.length} tuoi, oltre ai predefiniti` : "solo i predefiniti"}
          </span>
          <button className="btn ghost mini" onClick={() => apri("modelli")}>+ Nuovo modello</button>
          {((corrente && corrente.passi) || []).length > 0 && (
          <button className="btn ghost mini" onClick={() => chiedi({
            titolo: "Salvare questo percorso come modello?",
            testo: `Salvo la sequenza dei ${corrente ? corrente.passi.length : 0} passi di "${corrente ? corrente.nome : ""}" — nomi e tipo, senza i numeri — così puoi riusarla su altri clienti scegliendola dal modulo del funnel.`,
            etichetta: "Salva",
            campo: { label: "Nome del modello", valore: "" },
            azione: (nome) => upd("modelli", (l) => [
            { id: uid(), nome, creato: oggi(), passi: (corrente && corrente.passi || []).map((p) => [p.nome, p.tipo]) },
            ...l.filter((x) => x.nome.trim().toLowerCase() !== String(nome).trim().toLowerCase()),
            ]),
          })}>Salva questo come modello</button>
          )}
        </header>
        {(data.modelli || []).length > 0 ? (
          <div className="acts">
          {data.modelli.map((x) => (
            <div className="tkrow clic" key={x.id} onClick={() => apri("modelli", x)}>
            <div className="t">
              <b>{x.nome}</b>
              <div className="sub">
                {x.passi.map((p) => p[0]).join(" → ")}
                <span style={{ marginLeft: 8 }}>· salvato il {dataIt(x.creato)}</span>
              </div>
            </div>
            <button className="btn ghost mini" onClick={(e) => { e.stopPropagation(); chiedi({
              titolo: `Eliminare il modello "${x.nome}"?`,
              testo: "I funnel già creati con questo modello non vengono toccati: sparisce solo dall'elenco di quelli riusabili.",
              etichetta: "Elimina", pericolo: true,
              azione: () => upd("modelli", (l) => l.filter((y) => y.id !== x.id)),
            }); }}>Elimina</button>
            </div>
          ))}
          </div>
        ) : (
          <div className="nota" style={{ paddingTop: 14 }}>
          {corrente
            ? "Quando costruisci un percorso che funziona, salvalo da qui: lo ritrovi nel modulo del funnel, sotto i quattro predefiniti, e lo applichi a un altro cliente senza riscrivere i passi."
            : "Qui compaiono i percorsi che ti costruisci. Per crearne uno: aggiungi un funnel con “+ Aggiungi” in alto, componi i suoi passi, poi torna qui e premi “Salva questo come modello”. Da quel momento lo trovi nel modulo del funnel, sotto i quattro predefiniti, e lo riusi su qualsiasi cliente."}
          </div>
        )}
        </section>

    );

    const f = lista.find((x) => x.id === funnelAperto) || lista[0];
    const cli = f ? cliente(f.clienteId) : null;

    if (!lista.length) {
      return (
        <>
          <section className="card">
            <div className="empty">
              <b>Nessun funnel</b>
              Un funnel qui è la sequenza di passi dal primo contatto alla conversione, con i numeri di ciascuno:
              serve a vedere dove si perdono le persone e a generare la pagina di destinazione.
              Premi “+ Aggiungi” in alto per crearne uno: puoi partire da un modello o da zero.
            </div>
          </section>
          {schedaModelli(null)}
        </>
      );
    }

    const passi = (f.passi || []).map((p, i, arr) => {
      const tasso = p.visite ? p.conversioni / p.visite : 0;
      const perdita = Math.max(0, (p.visite || 0) - (p.conversioni || 0));
      const scollegato = i > 0 && arr[i - 1].conversioni && Math.abs(arr[i - 1].conversioni - p.visite) > arr[i - 1].conversioni * 0.05;
      return { ...p, tasso, perdita, scollegato };
    });
    const validi = passi.filter((p) => p.visite > 0);
    const peggiore = validi.length ? validi.slice(1).reduce((a, b) => (b.tasso < a.tasso ? b : a), validi[1] || validi[0]) : null;
    const finali = passi.length ? passi[passi.length - 1].conversioni || 0 : 0;
    const inizio = passi.length ? passi[0].visite || 0 : 0;
    const collegate = m.campagne.filter((k) => k.funnelId === f.id);
    const spesaReale = collegate.reduce((s, k) => s + k.d.sp, 0);
    const spesaUsata = spesaReale > 0 ? spesaReale : (f.spesa || 0);
    const costoConv = finali ? spesaUsata / finali : 0;
    const ricavo = finali * (f.valoreMedio || 0);
    const diag = peggiore ? DIAGNOSI_PASSO[peggiore.tipo] || DIAGNOSI_PASSO.pagina : null;

    return conValuta(valutaCliente(data, f.clienteId), () => (
      <>
        {lista.length > 1 && (
          <section className="card">
            <header>
              <div className="ctrl">
                <select value={f.id} onChange={(e) => setFunnelAperto(e.target.value)} aria-label="Scegli il funnel">
                  {lista.map((x) => <option key={x.id} value={x.id}>{x.nome}{!attivo ? ` · ${nomeCli(x.clienteId)}` : ""}</option>)}
                </select>
              </div>
            </header>
          </section>
        )}

        <div className="kpis">
          <div className="kpi"><IcoKpi t="Entrano" /><div className="eyebrow">Entrano</div><div className="v">{int(inizio)}</div><div className="s">al primo passo</div></div>
          <div className="kpi"><IcoKpi t="Arrivano in fondo" /><div className="eyebrow">Arrivano in fondo</div><div className="v">{int(finali)}</div><div className="s">conversione totale {inizio ? pctv((finali / inizio) * 100, 2) : "—"}</div></div>
          <div className="kpi"><IcoKpi t="Costo per conversione" /><div className="eyebrow">Costo per conversione</div><div className="v">{costoConv ? eur(costoConv) : "—"}</div><div className="s">{eur0(spesaUsata)} di spesa{spesaReale > 0 ? ` · da ${collegate.length} campagne collegate` : ""}</div></div>
          <div className="kpi"><IcoKpi t="Ritorno" /><div className="eyebrow">Ritorno</div><div className="v">{spesaUsata && ricavo ? `${fmt(ricavo / spesaUsata)}×` : "—"}</div><div className="s">{ricavo ? `${eur0(ricavo)} attesi` : "valore medio non impostato"}</div></div>
        </div>

        {(() => {
          /* Quello che il CRM già sa: impression e clic dalle campagne collegate,
             lead e appuntamenti dal ciclo di vendita. Il resto lo sa solo la pagina. */
          const coll = m.campagne.filter((k) => k.funnelId === f.id);
          const imp = coll.reduce((a, k) => a + k.d.imp, 0);
          const clk = coll.reduce((a, k) => a + (k.d.clk || k.d.cl), 0);
          const lead = coll.reduce((a, k) => a + k.d.ld, 0);
          const vz = data.vendite.filter((x) => x.clienteId === f.clienteId);
          const gz = data.giornate.filter((x) => x.clienteId === f.clienteId);
          const somma = (c) => (gz.length ? gz : vz).reduce((a, x) => a + (x[c] || 0), 0);
          const fissate = somma("demoFissate"), svolte = somma("demoSvolte");
          const chius = data.chiusure.filter((c) => c.clienteId === f.clienteId && coll.some((k) => k.id === c.campagnaId)).length;

          const riempibile = coll.length > 0 && (f.passi || []).length > 0;
          const riempi = () => {
            const nuovi = (f.passi || []).map((p, i, arr) => {
              const prec = i > 0 ? arr[i - 1] : null;
              const q = { ...p };
              if (p.tipo === "annuncio") { q.visite = imp; q.conversioni = clk; }
              else if (["pagina", "vsl"].includes(p.tipo)) { q.visite = prec && prec.tipo === "annuncio" ? clk : q.visite; }
              else if (p.tipo === "form") { q.conversioni = lead; }
              else if (p.tipo === "calendario") { q.visite = lead; if (fissate) q.conversioni = fissate; }
              else if (p.tipo === "chiamata") { if (fissate) q.visite = fissate; if (svolte) q.conversioni = svolte; }
              else if (["checkout", "upsell"].includes(p.tipo)) { if (svolte) q.visite = svolte; if (chius) q.conversioni = chius; }
              return q;
            });
            chiedi({
              titolo: "Riempire i passi con i dati del CRM?",
              testo: `Scrivo quello che so già: impression e clic sul link dalle ${coll.length} campagne collegate, i lead come moduli inviati, e — se li hai registrati — appuntamenti fissati e svolti dal ciclo di vendita.\n\nI passaggi che vivono solo sulla pagina, come quante persone dalla landing arrivano al modulo, restano come sono: quelli li prendi dalle statistiche del sito.`,
              etichetta: "Riempi",
              azione: () => upd("funnel", (l) => l.map((x) => (x.id === f.id ? { ...x, passi: nuovi } : x))),
            });
          };

          return riempibile ? (
            <div className="nota" style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
              <button className="btn ghost mini" onClick={riempi}>Riempi dai dati del CRM</button>
              <span>
                Dalle campagne collegate: {int(imp)} impression, {int(clk)} clic sul link{lead ? `, ${int(lead)} lead` : ""}.
                {fissate ? ` Dal ciclo di vendita: ${int(fissate)} appuntamenti fissati, ${int(svolte)} svolti.` : ""}
              </span>
            </div>
          ) : null;
        })()}

        {(() => {
          const suoi = m.campagne.filter((k) => k.clienteId === f.clienteId && k.archiviata !== "si")
            .sort((a, b) => (b.inizio || "").localeCompare(a.inizio || ""));
          if (!suoi.length) return null;
          const collegateId = suoi.filter((k) => k.funnelId === f.id).map((k) => k.id);
          const tutte = collegateId.length === suoi.length;
          const imposta = (ids, valore) => upd("campagne", (l) => l.map((k) => (ids.includes(k.id) ? { ...k, funnelId: valore } : k)));
          const spesaColl = suoi.filter((k) => collegateId.includes(k.id)).reduce((s2, k) => s2 + k.d.sp, 0);

          return (
            <section className="card">
              <header>
                <h2>Campagne collegate</h2>
                <span className="sp eyebrow">
                  {collegateId.length} di {suoi.length}{spesaColl ? ` · ${eur0(spesaColl)} di spesa` : ""}
                </span>
                <button className="btn ghost mini"
                  onClick={() => imposta(tutte ? collegateId : suoi.map((k) => k.id), tutte ? "" : f.id)}>
                  {tutte ? "Scollega tutte" : "Collega tutte"}
                </button>
                {suoi.some((k) => ["attiva", "test"].includes(k.stato) && k.funnelId !== f.id) && (
                  <button className="btn ghost mini"
                    onClick={() => imposta(suoi.filter((k) => ["attiva", "test"].includes(k.stato)).map((k) => k.id), f.id)}>
                    Solo le attive
                  </button>
                )}
              </header>
              <div className="sceltarigaintest larga" style={{ margin: "0 16px" }}>
                <span />
                <span>Campagna</span>
                <span>Stato</span>
                <span>Speso</span>
                <span>Lead</span>
                <span>Funnel attuale</span>
              </div>
              <div className="sceltalista" style={{ maxHeight: "none", padding: "0 16px 14px" }}>
                {suoi.map((k) => {
                  const mia = k.funnelId === f.id;
                  const altrove = k.funnelId && !mia ? (data.funnel.find((x) => x.id === k.funnelId) || {}).nome : "";
                  return (
                    <label className="sceltariga larga" key={k.id} data-on={mia ? "1" : "0"}>
                      <input type="checkbox" checked={mia} aria-label={`Collega ${k.nome}`}
                        onChange={(e) => imposta([k.id], e.target.checked ? f.id : "")} />
                      <span className="nm">{k.nome}</span>
                      <span className="st"><span className="tag" data-k={k.stato}>{etStato(k)}</span>{k.accountBloccato ? <div className="acc-bloccato" title={k.stato === "errore" ? "La campagna era accesa: l'account l'ha fermata" : "La campagna era già spenta: l'account non la ferma, ma se la riaccendi non erogherà"}>⚠ account {k.accountBloccato}{k.stato !== "errore" ? " · campagna già spenta" : ""}</div> : null}</span>
                      <span className="st mono">{eur0(k.d.sp)}</span>
                      <span className="st mono">{k.d.ld ? int(k.d.ld) : "—"}</span>
                      <span className="alt">{altrove ? `già su ${altrove}` : ""}</span>
                    </label>
                  );
                })}
              </div>
              <div className="nota">
                Le stesse caselle sono nel modulo del funnel, quando lo crei o lo modifichi.
                Una campagna appartiene a un funnel solo: spuntandola qui, se era collegata altrove viene spostata.
                Dal collegamento nascono il costo per conversione di questo funnel e la riga nei totali per funnel in Advertising.
              </div>
            </section>
          );
        })()}


        <section className="card">
          <header>
            <h2>{f.nome}</h2>
            <span className="sp eyebrow">{et(OBIETTIVI_FUNNEL, f.obiettivo)}</span>
            <button className="btn ghost mini" onClick={() => apri("funnel", f)}>Modifica</button>
          </header>
          {validi.length > 1 && (
            <div style={{ padding: "18px 16px 6px" }}>
              <Funnel steps={passi.slice(0, 4).map((p) => ({ k: p.nome, n: p.visite || 0, sub: et(TIPI_PASSO, p.tipo).toUpperCase() }))} />
            </div>
          )}
          <div className="wrap">
            <table className="tbl">
              <thead><tr><th>Passo</th><th>Tipo</th><th className="r">Entrano</th><th className="r">Passano</th><th className="r">Tasso</th><th className="r">Persi</th></tr></thead>
              <tbody>
                {passi.map((p) => (
                  <tr key={p.id}>
                    <td><b>{p.nome}</b>{p.scollegato && <div className="sub neg">chi entra qui non corrisponde a chi è passato prima</div>}</td>
                    <td className="sub">{et(TIPI_PASSO, p.tipo)}</td>
                    <td className="r">{int(p.visite)}</td>
                    <td className="r">{int(p.conversioni)}</td>
                    <td className={"r " + (peggiore && p.id === peggiore.id ? "neg" : "")}>{p.visite ? pctv(p.tasso * 100, 1) : "—"}</td>
                    <td className="r sub">{int(p.perdita)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {!passi.length && <div className="empty"><b>Nessun passo</b>Apri Modifica e scegli un modello, oppure costruisci la sequenza a mano.</div>}
        </section>

        {peggiore && diag && peggiore.visite > 0 && (
          <article className="al" data-g="attenzione">
            <div className="meta">
              <span className="tag">passaggio debole</span>
              <span className="misura">{peggiore.nome} · <b>{pctv(peggiore.tasso * 100, 1)}</b></span>
            </div>
            <h3>Qui perdi {int(peggiore.perdita)} persone</h3>
            <dl>
              <dt>Cosa succede</dt>
              <dd>Su {int(peggiore.visite)} che arrivano a "{peggiore.nome}" ne prosegue il {pctv(peggiore.tasso * 100, 1)}. È il passaggio con la resa più bassa del funnel: recuperare qualche punto qui vale più che aumentare il traffico a monte.</dd>
              <dt>Causa probabile</dt><dd>{diag.causa}</dd>
              <dt>Cosa fare</dt><dd>{diag.azione}</dd>
            </dl>
          </article>
        )}

        <section className="card">
          <header>
            <h2>Pagina di destinazione</h2>
            <span className="sp eyebrow">file HTML autonomo, da caricare dove vuoi</span>
            <button className="btn ghost mini" onClick={() => {
              const html = paginaHTML(f, cli, data.agenzia, cli && cli.logo ? cli.logo : logoAgenzia);
              const w = window.open("", "_blank");
              if (w) { w.document.write(html); w.document.close(); }
            }}>Anteprima</button>
            <button className="btn" onClick={() => scaricaFile(
              `${(f.nome || "pagina").toLowerCase().replace(/[^a-z0-9]+/g, "-")}.html`,
              paginaHTML(f, cli, data.agenzia, cli && cli.logo ? cli.logo : logoAgenzia)
            )}>Scarica la pagina</button>
          </header>
          <table className="tbl"><tbody>
            <tr><td className="sub">Titolo</td><td><b>{(f.pagina || {}).titolo || "—"}</b></td></tr>
            <tr><td className="sub">Sottotitolo</td><td>{(f.pagina || {}).sottotitolo || "—"}</td></tr>
            <tr><td className="sub">Pulsante</td><td>{(f.pagina || {}).cta || "—"}</td></tr>
            <tr><td className="sub">Modulo</td><td>{(f.pagina || {}).webhook ? "collegato a un webhook" : <span className="neg">nessun webhook: i dati non vanno da nessuna parte</span>}</td></tr>
          </tbody></table>
          <div className="nota">
            Il file generato è una pagina completa, con il logo del cliente e i colori del tuo brand: la carichi su un hosting,
            su Netlify o dentro il sito del cliente. Il modulo invia i dati al webhook che hai indicato, quindi da lì li porti
            dove vuoi. Quello che non può fare, e nessun file da solo può fare, è ospitarsi da sé o gestire un dominio.
          </div>
        </section>
      </>
    ));
  };

  /* ---------- 4. pianificatore a ritroso ---------- */
  const SezionePianificatore = () => {
    const t = sommaVendite(venditeVista);
    const r = {
      r1: t.lead ? t.discoveryFissate / t.lead : 0,
      r2: t.discoveryFissate ? t.discoverySvolte / t.discoveryFissate : 0,
      r3: t.discoverySvolte ? t.demoFissate / t.discoverySvolte : 0,
      r4: t.demoFissate ? t.demoSvolte / t.demoFissate : 0,
      r5: t.demoSvolte ? t.chiusure / t.demoSvolte : 0,
    };
    const cplStorico = t.lead ? t.spesa / t.lead : 0;
    const contrattoMedio = t.chiusure ? t.valoreContratto / t.chiusure : 0;
    const completo = Object.values(r).every((x) => x > 0);
    const cpl = parseNum(piano.cpl) || cplStorico;
    const ob = parseNum(piano.obiettivo) || 0;

    const demoSvolte = r.r5 ? ob / r.r5 : 0;
    const demoFissate = r.r4 ? demoSvolte / r.r4 : 0;
    const discoverySvolte = r.r3 ? demoFissate / r.r3 : 0;
    const discoveryFissate = r.r2 ? discoverySvolte / r.r2 : 0;
    const lead = r.r1 ? discoveryFissate / r.r1 : 0;
    const budget = lead * cpl;
    const ricavo = ob * contrattoMedio;

    const passi = [
      { l: "Lead necessari", v: lead, n: `${fmt(r.r1 * 100, 1)}% diventa discovery fissata` },
      { l: "Discovery fissate", v: discoveryFissate, n: `show up ${fmt(r.r2 * 100, 1)}%` },
      { l: "Discovery svolte", v: discoverySvolte, n: `${fmt(r.r3 * 100, 1)}% arriva alla demo` },
      { l: "Demo fissate", v: demoFissate, n: `show up ${fmt(r.r4 * 100, 1)}%` },
      { l: "Demo svolte", v: demoSvolte, n: `chiusura ${fmt(r.r5 * 100, 1)}%` },
      { l: "Chiusure", v: ob, n: "l'obiettivo che hai messo" },
    ];

    return conValuta(valutaComune(venditeVista.map((v) => v.clienteId)), () => (
      <>
        <section className="card">
          <header><h2>Obiettivo</h2><span className="sp eyebrow">i tassi arrivano dal tuo storico</span></header>
          <div className="fields" style={{ gridTemplateColumns: "1fr 1fr 1fr" }}>
            <div className="f">
              <label htmlFor="pi-ob">Chiusure da fare</label>
              <input id="pi-ob" className="num" type="text" inputMode="decimal" value={piano.obiettivo}
                onChange={(e) => setPiano({ ...piano, obiettivo: e.target.value.replace(/[^\d.,]/g, "") })} />
            </div>
            <div className="f">
              <label htmlFor="pi-cpl">Costo per lead</label>
              <input id="pi-cpl" className="num" type="text" inputMode="decimal" placeholder={perInput(Math.round(cplStorico * 100) / 100)}
                value={piano.cpl} onChange={(e) => setPiano({ ...piano, cpl: e.target.value.replace(/[^\d.,]/g, "") })} />
              <span className="hint">storico {cplStorico ? eur(cplStorico) : "non disponibile"}</span>
            </div>
            <div className="f">
              <label>Contratto medio</label>
              <input value={contrattoMedio ? eur(contrattoMedio) : "—"} readOnly />
              <span className="hint">dalle chiusure registrate</span>
            </div>
          </div>
        </section>

        {!completo || !ob ? (
          <section className="card">
            <div className="empty">
              <b>{!ob ? "Metti un obiettivo" : "Servono più dati commerciali"}</b>
              {!ob ? "Scrivi quante chiusure vuoi fare e vedi cosa serve a monte."
                : "Il calcolo a ritroso ha bisogno di almeno un mese con lead, discovery e demo compilati."}
            </div>
          </section>
        ) : (
          <>
            <div className="kpis">
              <div className="kpi"><IcoKpi t="Budget necessario" /><div className="eyebrow">Budget necessario</div><div className="v">{eur0(budget)}</div><div className="s">a {eur(cpl)} per lead</div></div>
              <div className="kpi"><IcoKpi t="Lead da generare" /><div className="eyebrow">Lead da generare</div><div className="v">{int(lead)}</div><div className="s">{int(Math.ceil(lead / 30))} al giorno per un mese</div></div>
              <div className="kpi"><IcoKpi t="Ricavo atteso" /><div className="eyebrow">Ricavo atteso</div><div className="v">{eur0(ricavo)}</div><div className="s">contratto, non incassato</div></div>
              <div className="kpi"><IcoKpi t="Ritorno sulla spesa" /><div className="eyebrow">Ritorno sulla spesa</div><div className="v">{budget ? `${fmt(ricavo / budget)}×` : "—"}</div><div className="s">{attivo && roasPareggio(attivo) ? `pareggio a ${fmt(roasPareggio(attivo))}×` : "margine non impostato"}</div></div>
            </div>

            <section className="card">
              <header><h2>Cosa serve a monte</h2><span className="sp eyebrow">dal basso verso l'alto</span></header>
              <table className="tbl">
                <thead><tr><th>Passaggio</th><th className="r">Quantità</th><th>Tasso applicato</th></tr></thead>
                <tbody>
                  {passi.map((p) => (
                    <tr key={p.l}><td><b>{p.l}</b></td><td className="r">{int(Math.ceil(p.v))}</td><td className="sub">{p.n}</td></tr>
                  ))}
                </tbody>
              </table>
              <div className="nota">
                Con questi numeri servono circa {int(Math.ceil(discoverySvolte))} discovery e {int(Math.ceil(demoSvolte))} demo da tenere davvero:
                verifica che la capacità commerciale ci arrivi, altrimenti il budget in più produce solo lead non lavorati.
              </div>
            </section>
          </>
        )}
      </>
    ));
  };

  /* ---------- 6 e 7. diario ed esperimenti ---------- */
  const SezioneDiario = () => {
    const note = [...(attivo ? S.annotazioni : data.annotazioni)].sort((a, b) => (a.data < b.data ? 1 : -1));
    const exp = [...(attivo ? S.esperimenti : data.esperimenti)].sort((a, b) => (a.inizio < b.inizio ? 1 : -1));
    const nomeCampagna = (id) => (data.campagne.find((k) => k.id === id) || {}).nome || "";
    return (
      <>
        <section className="card">
          <header><h2>Diario delle decisioni</h2><span className="sp" />
            <button className="btn ghost mini" onClick={() => apri("annotazioni")}>+ Annotazione</button>
          </header>
          <div className="acts">
            {note.map((n) => (
              <div className="tkrow clic" key={n.id} onClick={() => apri("annotazioni", n)}>
                <div className="t">
                  <b>{n.testo}</b>
                  <div className="meta" style={{ marginTop: 5 }}>
                    <span className="tag" data-k="qualificato">{et(TIPI_NOTA, n.tipo)}</span>
                    {!attivo && <span className="pill">{nomeCli(n.clienteId)}</span>}
                    {n.campagnaId && <span className="pill">{nomeCampagna(n.campagnaId)}</span>}
                  </div>
                </div>
                <span className="pill">{dataIt(n.data)}</span>
              </div>
            ))}
            {!note.length && (
              <div className="empty"><b>Diario vuoto</b>Segna ogni volta che cambi budget, creatività o pubblico: fra due mesi saprai perché la curva ha svoltato.</div>
            )}
          </div>
        </section>

        <section className="card">
          <header><h2>Esperimenti</h2><span className="sp" />
            <button className="btn ghost mini" onClick={() => apri("esperimenti")}>+ Esperimento</button>
          </header>
          <div className="wrap">
            <table className="tbl">
              <thead><tr><th>Ipotesi</th><th>Variabile</th><th>Esito</th><th>Risultato</th><th className="r">Periodo</th></tr></thead>
              <tbody>
                {exp.map((e) => (
                  <tr key={e.id} className="clic" onClick={() => apri("esperimenti", e)}>
                    <td><b>{e.ipotesi}</b>{e.conclusione && <div className="sub">{e.conclusione}</div>}</td>
                    <td className="sub">{e.variabile}<div className="sub">misurato su {e.metrica || "—"}</div></td>
                    <td><span className="tag" data-k={e.stato === "vinto" ? "attivo" : e.stato === "corso" ? "test" : ""}>{et(STATI_ESPERIMENTO, e.stato)}</span></td>
                    <td className="sub">{e.risultato || "—"}</td>
                    <td className="r sub">{dataIt(e.inizio)}{e.fine ? ` → ${dataIt(e.fine)}` : ""}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {!exp.length && (
            <div className="empty"><b>Nessun esperimento</b>Un test scritto vale dieci intuizioni ricordate male: ipotesi, variabile, esito, conclusione.</div>
          )}
        </section>
      </>
    );
  };

  /* ---------- 5. report per il cliente ---------- */
  const SezioneReport = () => {
    const mm = mese(oggi());
    const camp = S.campagne.filter((k) => ["attiva", "test", "conclusa"].includes(k.stato));
    const T = camp.reduce((s, k) => ({ sp: s.sp + k.d.sp, cv: s.cv + k.d.cv, rc: s.rc + k.d.rc, cl: s.cl + k.d.cl, imp: s.imp + k.d.imp }), { sp: 0, cv: 0, rc: 0, cl: 0, imp: 0 });
    const chiusureMese = S.chiusure.filter((c) => mese(c.data || "") === mm);
    const fatti = S.task.filter((t) => t.stato === "fatto");
    const prossimi = ordinaTask(S.task.filter((t) => t.stato !== "fatto")).slice(0, 6);
    const noteMese = S.annotazioni.filter((n) => mese(n.data || "") === mm);
    const p = profitto(T.rc, T.sp, attivo);

    return (
      <>
        <section className="card nostampa">
          <header>
            <h2>Report di {meseIt(mm)}</h2>
            <span className="sp eyebrow">generato dai dati già inseriti</span>
            <button className="btn" onClick={() => window.print()}>Stampa o salva in PDF</button>
          </header>
          <div className="nota" style={{ paddingTop: 14 }}>
            La stampa contiene solo il report qui sotto: menu, filtri e pulsanti restano fuori.
          </div>
        </section>

        <div className="report">
          <div className="rtesta">
            <div className="riga">
              <Logo src={attivo.logo} nome={attivo.nome} size={54} />
              <div>
                <h1>{attivo.nome}</h1>
                <div className="sub">Report advertising · {meseIt(mm)}</div>
              </div>
            </div>
            <div className="riga" style={{ marginLeft: "auto" }}>
              <div style={{ textAlign: "right" }}>
                <div className="eyebrow">a cura di</div>
                <div style={{ fontWeight: 650 }}>{data.agenzia.nome}</div>
              </div>
              <Logo src={logoAgenzia} nome={data.agenzia.nome} size={44} />
            </div>
          </div>

          <div className="kpis">
            <div className="kpi"><IcoKpi t="Investito" /><div className="eyebrow">Investito</div><div className="v">{eur0(T.sp)}</div><div className="s">{camp.length} campagne</div></div>
            <div className="kpi"><IcoKpi t="Risultati" /><div className="eyebrow">Risultati</div><div className="v">{int(T.cv)}</div><div className="s">conversioni · CPA {T.cv ? eur(T.sp / T.cv) : "—"}</div></div>
            <div className="kpi"><IcoKpi /><div className="eyebrow">{T.rc ? "ROAS" : "Contratti chiusi"}</div><div className="v">{T.rc ? `${fmt(T.sp ? T.rc / T.sp : 0)}×` : eur0(chiusureMese.reduce((s, c) => s + (c.valoreContratto || 0), 0))}</div><div className="s">{T.rc ? `${eur0(T.rc)} di ricavi` : `${chiusureMese.length} chiusure nel mese`}</div></div>
            <div className="kpi"><IcoKpi t="Margine stimato" /><div className="eyebrow">Margine stimato</div><div className="v">{margineNetto(attivo) ? eur0(p) : "—"}</div><div className="s">{margineNetto(attivo) ? `al netto di margine ${fmt(attivo.margineLordo, 0)}% e resi` : "margine non impostato"}</div></div>
          </div>

          <section className="card">
            <header><h2>Le campagne</h2></header>
            <div className="wrap"><table className="tbl">
              <thead><tr><th>Campagna</th><th className="r">Speso</th><th className="r">CTR link</th><th className="r">Lead</th><th className="r">CPL</th><th className="r">Conv.</th><th className="r">CPA</th><th className="r">ROAS</th></tr></thead>
              <tbody>
                {camp.map((k) => (
                  <tr key={k.id}>
                    <td><b>{k.nome}</b><div className="sub">{et(PIATTAFORME, k.piattaforma)} · {et(OBIETTIVI, k.obiettivo)}</div></td>
                    <td className="r">{eur(k.d.sp)}</td>
                    <td className="r">{pctv(k.d.ctrLink || k.d.ctr)}</td>
                    <td className="r">{k.d.ld ? int(k.d.ld) : "—"}</td>
                    <td className="r">{k.d.ld ? eur(k.d.cpl) : "—"}</td>
                    <td className="r">{int(k.d.cv)}</td>
                    <td className="r">{k.d.cv ? eur(k.d.cpa) : "—"}</td>
                    <td className="r">{k.d.rc ? `${fmt(k.d.roas)}×` : "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table></div>
          </section>

          {S.vendite.length > 0 && (
            <section className="card">
              <header><h2>Il percorso commerciale</h2></header>
              <Barre dati={FASI_VENDITA.map((f) => ({ k: f.l, v: sommaVendite(venditeVista)[f.k] || 0 }))} fmtv={int} />
            </section>
          )}

          <div className="two">
            <section className="card">
              <header><h2>Cosa ho fatto</h2></header>
              <div className="acts">
                {fatti.slice(0, 8).map((t) => (
                  <div className="tkrow" key={t.id}><div className="t"><b>{t.titolo}</b><div className="sub">{nomeArea(t.areaId)}</div></div></div>
                ))}
                {noteMese.map((n) => (
                  <div className="tkrow" key={n.id}><div className="t"><b>{n.testo}</b><div className="sub">{et(TIPI_NOTA, n.tipo)} · {dataIt(n.data)}</div></div></div>
                ))}
                {!fatti.length && !noteMese.length && <div className="empty">Nessuna attività conclusa registrata.</div>}
              </div>
            </section>
            <section className="card">
              <header><h2>Cosa faccio adesso</h2></header>
              <div className="acts">
                {prossimi.map((t) => (
                  <div className="tkrow" key={t.id}><div className="t"><b>{t.titolo}</b><div className="sub">{nomeArea(t.areaId)} · entro {dataIt(t.scadenza)}</div></div></div>
                ))}
                {!prossimi.length && <div className="empty">Nessun task aperto.</div>}
              </div>
            </section>
          </div>
        </div>
      </>
    );
  };

  /* ---------- 8. economia dell'agenzia ---------- */
  const SezioneEconomia = () => {
    const costoOra = data.agenzia.costoOrario || 0;
    const da = piu(-30);
    const righe = data.clienti.filter((z) => z.stato !== "chiuso" && z.archiviato !== "si").map((z) => {
      const ore = data.task.filter((t) => t.clienteId === z.id && t.stato === "fatto" && (t.scadenza || "") >= da)
        .reduce((s, t) => s + (t.stima || 0), 0);
      const costo = ore * costoOra;
      return { ...z, ore, costo, margine: (z.fee || 0) - costo, aperti: data.task.filter((t) => t.clienteId === z.id && t.stato !== "fatto").length };
    });
    const feeTot = righe.reduce((s, z) => s + (z.fee || 0), 0);
    const oreTot = righe.reduce((s, z) => s + z.ore, 0);
    const margineTot = righe.reduce((s, z) => s + z.margine, 0);
    const conc = feeTot ? Math.max(...righe.map((z) => (z.fee || 0) / feeTot), 0) : 0;
    const primo = righe.slice().sort((a, b) => (b.fee || 0) - (a.fee || 0))[0];

    /* Qui si parla di compensi: valuta del compenso, non quella delle campagne */
    const valuteFee = [...new Set(righe.map((z) => valutaCompenso(data, z.id)))];
    return conValuta(valuteFee.length === 1 ? valuteFee[0] : VALUTA, () => (
      <>
        <div className="kpis">
          <div className="kpi"><IcoKpi t="Ricorrente mensile" /><div className="eyebrow">Ricorrente mensile</div><div className="v">{eur0(feeTot)}</div><div className="s">{righe.length} clienti attivi</div></div>
          <div className="kpi"><IcoKpi t="Ore ultimi 30 giorni" /><div className="eyebrow">Ore ultimi 30 giorni</div><div className="v">{fmt(oreTot, 1)}</div><div className="s">{costoOra ? `valorizzate a ${eur(costoOra)}/h` : "imposta il costo orario"}</div></div>
          <div className="kpi"><IcoKpi t="Margine stimato" /><div className="eyebrow">Margine stimato</div><div className="v" style={{ color: margineTot < 0 ? "var(--acc)" : "var(--ink)" }}>{eur0(margineTot)}</div><div className="s">compensi meno tempo speso</div></div>
          <div className="kpi"><IcoKpi t="Concentrazione" /><div className="eyebrow">Concentrazione</div><div className="v" style={{ color: conc > 0.35 ? "var(--acc)" : "var(--ink)" }}>{pctv(conc * 100, 0)}</div><div className="s">{primo ? `quota di ${primo.nome}` : "—"}</div></div>
        </div>

        {conc > 0.35 && primo && (
          <article className="al" data-g="attenzione">
            <div className="meta"><span className="tag">rischio</span><span className="misura">quota <b>{pctv(conc * 100, 0)}</b> · soglia 35%</span></div>
            <h3>Troppo fatturato su un cliente solo</h3>
            <dl>
              <dt>Cosa succede</dt><dd>{primo.nome} vale il {pctv(conc * 100, 0)} del tuo ricorrente: se se ne va, non perdi un cliente, perdi lo stipendio di diversi mesi.</dd>
              <dt>Causa probabile</dt><dd>Crescita su un account che funziona bene, mentre l'acquisizione di nuovi clienti è ferma perché sei occupato a servirlo.</dd>
              <dt>Cosa fare</dt><dd>Tieni un blocco di ore a settimana per la tua acquisizione, come fosse un cliente pagante. È l'unico modo per non trattare da posizione debole al prossimo rinnovo.</dd>
            </dl>
          </article>
        )}

        <section className="card">
          <header><h2>Conto economico per cliente</h2><span className="sp eyebrow">ore dai task conclusi negli ultimi 30 giorni · ▸ apre i sottoclienti</span></header>
          <div className="wrap">
            <table className="tbl">
              <thead><tr><th>Cliente</th><th className="r">Compenso</th><th className="r">Ore</th><th className="r">Costo tempo</th><th className="r">Margine</th><th className="r">Margine %</th><th className="r">Quota</th></tr></thead>
              <tbody>
                {mapCon(righeGruppi("economia", (ids) => {
                  const l = righe.filter((z) => ids.includes(z.id));
                  return l.length ? ["fee", "ore", "costo", "margine", "aperti"].reduce((o, k) => ({ ...o, [k]: l.reduce((x, z) => x + (z[k] || 0), 0) }), {}) : null;
                }, (a, b) => b.margine - a.margine), (r) => valutaCompenso(data, r.z.id), (z) => (
                  <tr key={z.key} className="clic" onClick={() => cambiaAccount(z.z.id)}>
                    <td style={{ paddingLeft: 14 + z.liv * 18 }}><div className="riga">
                      {z.liv === 0 && frecciaCli("economia", z.z.id, z.nSotto)}
                      {z.liv > 0 && <span className="sub">↳</span>}
                      <Logo src={z.z.logo} nome={z.z.nome} size={z.liv ? 22 : 28} />
                      <div><b>{z.z.nome}</b><div className="sub">
                        {z.solo ? "solo il cliente principale · " : z.liv === 0 && z.nSotto ? `con ${int(z.nSotto)} ${z.nSotto === 1 ? "sottocliente" : "sottoclienti"} · ` : ""}{z.aperti} task aperti
                      </div></div>
                    </div></td>
                    <td className="r">{eur(z.fee)}</td>
                    <td className="r">{fmt(z.ore, 1)}</td>
                    <td className="r">{eur(z.costo)}</td>
                    <td className={"r " + (z.margine < 0 ? "neg" : "pos")}>{eur(z.margine)}</td>
                    <td className={"r " + (z.margine < 0 ? "neg" : "")}>{z.fee ? pctv((z.margine / z.fee) * 100, 0) : "—"}</td>
                    <td className="r">{feeTot ? pctv(((z.fee || 0) / feeTot) * 100, 0) : "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="nota">
            Le ore sono quelle che hai stimato sui task: se non compili la stima, il margine risulta più alto del vero.
            {!costoOra && " Imposta il costo orario in Impostazioni per vedere i costi."}
          </div>
        </section>
      </>
    ));
  };

  /* ---------- Winner ads del cliente ---------- */
  /* Markdown essenziale delle analisi: titoli, elenchi e grassetto */
  const testoAnalisi = (t) => String(t || "").split("\n").map((riga, i) => {
    const grassetto = (x) => x.split(/(\*\*[^*]+\*\*)/g).map((p, j) => (/^\*\*[^*]+\*\*$/.test(p) ? <b key={j}>{p.slice(2, -2)}</b> : p));
    if (/^#{1,4}\s/.test(riga)) return <div key={i} style={{ fontWeight: 650, marginTop: 10 }}>{grassetto(riga.replace(/^#{1,4}\s/, ""))}</div>;
    if (!riga.trim()) return <div key={i} style={{ height: 8 }} />;
    return <div key={i} style={/^\s*([-*•]|\d+[.)])\s/.test(riga) ? { paddingLeft: 14 } : undefined}>{grassetto(riga)}</div>;
  });
  /* ---------- Academy: la piattaforma di corsi collegata al CRM ---------- */
  const academyCorsi = () => ((data.academy.corsi || []).length ? data.academy.corsi : ACADEMY_ESEMPIO);
  const lezioneLibera = (c, l) => !!(c.gratuito || c.sbloccato || l.gratuita);
  /* Lezioni da suggerire in una sezione del CRM (prima quelle libere) */
  const academyPer = (sez) => [
    ...academyCorsi().filter((c) => !c.esempio && c.sezione === sez).map((c) => ({ c, l: null })),
    ...academyCorsi().flatMap((c) => (c.lezioni || []).filter((l) => l.sezione === sez).map((l) => ({ c, l })))
      .sort((a, b) => lezioneLibera(b.c, b.l) - lezioneLibera(a.c, a.l)),
  ];
  const SuggerimentoAcademy = ({ sez }) => {
    const s2 = academyPer(sez)[0];
    if (!s2) return null;
    const libera = s2.l ? lezioneLibera(s2.c, s2.l) : false;
    return (
      <div className="meta-msg" style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", borderRadius: 10, border: "1px solid var(--line)", background: "var(--surf)" }}>
        {s2.c.copertina ? <img src={s2.c.copertina} alt="" aria-hidden="true" style={{ width: 64, height: 36, borderRadius: 5, objectFit: "cover", flex: "none" }} />
          : <img src={LOGO_ACADEMY} alt="" aria-hidden="true" style={{ width: 26, height: 26, flex: "none" }} />}
        {s2.l ? <span><b>Dall'Academy</b> · {s2.l.titolo} <span className="sub">({s2.c.titolo} · {s2.l.durata}{libera ? " · gratuita" : ""})</span></span>
          : <span><b>Dall'Academy</b> · {s2.c.titolo} <span className="sub">({s2.c.livello || "corso"})</span></span>}
        <span className="sp" />
        <button className="btn ghost mini" onClick={() => { setVista("academy"); setAcCorso(s2.c.id); setAcLezione(s2.l ? { corso: s2.c.id, lezione: s2.l.id } : null); }}>{s2.l ? (libera ? "Guarda" : "🔒 Scopri") : "Scopri il corso"}</button>
      </div>
    );
  };
  const SezioneAcademy = () => {
    const ac = data.academy;
    const cambia = (campi) => setData((d) => ({ ...d, academy: { ...d.academy, ...campi } }));
    const urlOk = (u) => /^https:\/\//i.test(String(u || "").trim());
    const esempio = !(ac.corsi || []).length;
    const corsi = academyCorsi();
    const viste = ac.viste || {};
    const segna = (id, v) => cambia({ viste: { ...viste, [id]: v } });
    const nLez = (c) => (c.lezioni || []).length;
    const nViste = (c) => (c.lezioni || []).filter((l) => viste[l.id]).length;
    const perc = (c) => (nLez(c) ? (nViste(c) / nLez(c)) * 100 : 0);
    const sblocca = () => (urlOk(ac.urlAcquisto) ? window.open(ac.urlAcquisto, "_blank", "noopener")
      : urlOk(ac.url) ? window.open(ac.url, "_blank", "noopener")
      : window.alert("Qui si aprirà la pagina di acquisto dell'Accademia: indicala nelle impostazioni dell'Academy, in fondo alla pagina."));
    const sfondo = (c, k = 0) => ({ background: `radial-gradient(120% 140% at ${85 - (k % 3) * 20}% ${10 + (k % 2) * 30}%, ${c.colore || "#8E2039"} 0%, #2A0E16 45%, #08080A 100%)` });
    const apriCorso = (c, l = null) => { setAcCorso(c.id); setAcLezione(l ? { corso: c.id, lezione: l.id } : null); };
    const primaLibera = (c) => (c.lezioni || []).find((l) => lezioneLibera(c, l));
    // In evidenza: il primo corso con contenuti gratuiti non ancora finito, altrimenti il primo
    const evid = corsi.find((c) => c.copertina && !c.esempio && perc(c) < 100) || corsi.find((c) => primaLibera(c) && perc(c) < 100) || corsi[0];
    const corso = corsi.find((c) => c.id === acCorso);
    const lezione = corso && acLezione && (corso.lezioni || []).find((l) => l.id === acLezione.lezione);
    const scorri = (e, dir) => { const r = e.currentTarget.closest(".nf-riga").querySelector(".nf-scorri"); r.scrollBy({ left: dir * r.clientWidth * 0.9, behavior: "smooth" }); };
    const Bollo = ({ c }) => (c.esempio ? <span className="bollo">esempio</span>
      : c.gratuito ? <span className="bollo gratis">gratis</span> : c.sbloccato ? <span className="bollo">tuo</span>
      : c.prezzo ? <span className="bollo chiuso">🔒 {c.prezzo}</span> : null);
    const CartaCorso = ({ c, k }) => (
      <div className="nf-card" role="button" tabIndex={0} onClick={() => apriCorso(c)} onKeyDown={(e) => { if (e.key === "Enter") apriCorso(c); }} aria-label={c.titolo}>
        {c.copertina ? <img className="cop-img" src={c.copertina} alt="" loading="lazy" /> : (
          <>
            <div className="cop" style={sfondo(c, k)} />
            <span className="grande" aria-hidden="true">{(c.titolo || "A")[0]}</span>
            <img className="logo-n" src={LOGO_ACADEMY} alt="" aria-hidden="true" />
            <div className="tit">{c.titolo}<small>{c.livello ? `${c.livello} · ` : ""}{nLez(c) ? `${nLez(c)} lezioni` : "corso"}</small></div>
          </>
        )}
        <Bollo c={c} />
        {nViste(c) > 0 && <div className="avanz"><i style={{ width: `${perc(c)}%` }} /></div>}
      </div>
    );
    const CartaLezione = ({ c, l, k }) => (
      <div className="nf-card" role="button" tabIndex={0} onClick={() => apriCorso(c, l)} onKeyDown={(e) => { if (e.key === "Enter") apriCorso(c, l); }} aria-label={l.titolo}>
        <div className="cop" style={sfondo(c, k + 1)} />
        <span className="grande" aria-hidden="true">{(c.lezioni || []).indexOf(l) + 1}</span>
        <span className="bollo gratis">gratis</span>
        <div className="tit">{l.titolo}<small>{c.titolo} · {l.durata}</small></div>
        {viste[l.id] && <div className="avanz"><i style={{ width: "100%" }} /></div>}
      </div>
    );
    const Riga = ({ titolo, nota, apri, children }) => (
      <div className="nf-riga">
        <h3>{titolo}{nota && <small>{nota}</small>}{apri && <button className="nf-tutti" onClick={apri}>Vedi tutti ›</button>}</h3>
        <button className="nf-freccia sx" aria-label="Indietro" onClick={(e) => scorri(e, -1)}>‹</button>
        <div className="nf-scorri">{children}</div>
        <button className="nf-freccia dx" aria-label="Avanti" onClick={(e) => scorri(e, 1)}>›</button>
      </div>
    );
    const inCorso = corsi.filter((c) => nViste(c) > 0 && perc(c) < 100);
    const scheda = ACADEMY_MACRO.some((m2) => m2.v === acFiltro) ? acFiltro : "home";
    const diMacro = (v) => corsi.filter((c) => (c.macro || "corsi") === v);
    const gratis = corsi.flatMap((c) => (c.lezioni || []).filter((l) => lezioneLibera(c, l) && !c.sbloccato).map((l) => ({ c, l })));
    const miei = corsi.filter((c) => c.sbloccato);
    return (
      <div className="nf">
        {esempio && <div className="meta-msg" style={{ borderRadius: 10, border: "1px dashed var(--acc-b)" }}><b>Anteprima.</b> I corsi della sezione Corsi hanno titoli e copertine veri; descrizioni e lezioni, e i contenuti segnati «esempio» nelle altre sezioni, arriveranno dalla piattaforma.</div>}

        <nav className="nf-schede" aria-label="Sezioni dell'Academy">
          {[{ v: "home", l: "Home" }, ...ACADEMY_MACRO].map((m2) => (
            <button key={m2.v} aria-current={scheda === m2.v ? "page" : undefined} data-on={scheda === m2.v ? "1" : "0"} onClick={() => setAcFiltro(m2.v)}>
              {m2.l}{m2.v !== "home" && <small>{diMacro(m2.v).length}</small>}
            </button>
          ))}
        </nav>

        {scheda === "home" && evid && (
          <section className="nf-hero">
            <div className="sfondo" style={sfondo(evid, 2)} />
            {evid.copertina ? <img src={evid.copertina} alt="" aria-hidden="true" className="nf-hero-cop" /> : <img src={LOGO_ACADEMY} alt="" aria-hidden="true" className="nf-sigillo" />}
            <div>
              <div className="marchio"><img src={LOGO_ACADEMY} alt="" aria-hidden="true" style={{ width: 22, height: 22 }} />{ac.nome || "Accademia dell'Advertising"}</div>
              <h1>{evid.titolo}</h1>
              <div className="meta">
                {evid.gratuito ? <span className="tag" data-k="attivo">gratis</span> : evid.prezzo ? <span className="tag" data-k="alta">{evid.prezzo}</span> : null}
                <span>{(ACADEMY_MACRO.find((m2) => m2.v === (evid.macro || "corsi")) || {}).l}</span><span>·</span><span>{evid.livello}</span>
                {nLez(evid) > 0 && <><span>·</span><span>{nLez(evid)} lezioni</span></>}
                {nViste(evid) > 0 && <><span>·</span><span>seguite {nViste(evid)}</span></>}
              </div>
              <p>{evid.descrizione}</p>
              <div className="riga" style={{ gap: 10, flexWrap: "wrap" }}>
                {primaLibera(evid) && <button className="nf-play" onClick={() => apriCorso(evid, primaLibera(evid))}>▶ {nViste(evid) ? "Riprendi" : evid.gratuito ? "Guarda" : "Guarda l'anteprima"}</button>}
                <button className="nf-info" onClick={() => apriCorso(evid)}>ⓘ Altre informazioni</button>
                {!evid.gratuito && !evid.sbloccato && <button className="nf-sblocca" onClick={sblocca}>{evid.prezzo ? `Sblocca · ${evid.prezzo}` : "Accedi al corso"}</button>}
              </div>
            </div>
          </section>
        )}

        {scheda === "home" ? (
          <>
            {inCorso.length > 0 && <Riga titolo="Continua a guardare">{inCorso.map((c, k) => <CartaCorso key={c.id} c={c} k={k} />)}</Riga>}
            {miei.length > 0 && <Riga titolo="I tuoi corsi">{miei.map((c, k) => <CartaCorso key={c.id} c={c} k={k} />)}</Riga>}
            {ACADEMY_MACRO.map((m2, mi) => diMacro(m2.v).length > 0 && (
              <Riga key={m2.v} titolo={m2.l} nota={String(diMacro(m2.v).length)} apri={() => setAcFiltro(m2.v)}>{diMacro(m2.v).map((c, k) => <CartaCorso key={c.id} c={c} k={k + mi} />)}</Riga>
            ))}
            {gratis.length > 0 && <Riga titolo="Gratis per te" nota={`${gratis.length} lezioni libere`}>{gratis.map(({ c, l }, k) => <CartaLezione key={l.id} c={c} l={l} k={k} />)}</Riga>}
          </>
        ) : (
          <div style={{ display: "grid", gap: 14 }}>
            <div><h2 style={{ margin: "0 0 4px", fontSize: 26, letterSpacing: "-.02em" }}>{(ACADEMY_MACRO.find((m2) => m2.v === scheda) || {}).l}</h2>
              <div className="sub">{(ACADEMY_MACRO.find((m2) => m2.v === scheda) || {}).d}</div></div>
            <div className="nf-griglia">{diMacro(scheda).map((c, k) => <CartaCorso key={c.id} c={c} k={k} />)}</div>
            {!diMacro(scheda).length && <section className="card"><div className="empty"><b>Ancora niente qui</b>I contenuti di questa sezione arriveranno dalla piattaforma.</div></section>}
          </div>
        )}

        <section className="card">
          <header><h2>Impostazioni dell'Academy</h2><span className="sp eyebrow">collegamento con la piattaforma</span>
            {urlOk(ac.url) && <a className="btn ghost mini" href={ac.url} target="_blank" rel="noopener noreferrer" style={{ textDecoration: "none" }}>Apri la piattaforma ↗</a>}</header>
          <div className="riga" style={{ gap: 14, flexWrap: "wrap", padding: "14px 18px", alignItems: "flex-end" }}>
            <label style={{ display: "grid", gap: 4, flex: "1 1 220px" }}>
              <span className="eyebrow">Nome</span>
              <input defaultValue={ac.nome} onBlur={(e) => cambia({ nome: e.target.value.trim() || "Accademia dell'Advertising" })} />
            </label>
            <label style={{ display: "grid", gap: 4, flex: "2 1 280px" }}>
              <span className="eyebrow">Indirizzo della piattaforma</span>
              <input defaultValue={ac.url} placeholder="https://… (area corsi)" onBlur={(e) => cambia({ url: e.target.value.trim() })} />
            </label>
            <label style={{ display: "grid", gap: 4, flex: "2 1 280px" }}>
              <span className="eyebrow">Pagina di acquisto</span>
              <input defaultValue={ac.urlAcquisto || ""} placeholder="https://… (dove si paga l'accesso)" onBlur={(e) => cambia({ urlAcquisto: e.target.value.trim() })} />
            </label>
          </div>
          {[ac.url, ac.urlAcquisto].some((u) => u && !urlOk(u)) && <div className="meta-msg" data-tipo="errore">Gli indirizzi devono iniziare con https://</div>}
        </section>

        {corso && (
          <div className="nf-scrim" onClick={(e) => { if (e.target === e.currentTarget) { setAcLezione(null); setAcCorso(""); } }}>
            <div className="nf-modal" role="dialog" aria-modal="true" aria-label={corso.titolo}>
              {lezione ? (
                <div style={{ position: "relative", background: "#000" }}>
                  <button className="nf-chiudi" aria-label="Chiudi la lezione" onClick={() => setAcLezione(null)}>×</button>
                  {lezioneLibera(corso, lezione) ? (
                    /^https:\/\//.test(lezione.video || "") ? (
                      <iframe title={lezione.titolo} src={lezione.video} style={{ width: "100%", aspectRatio: "16 / 9", border: 0, display: "block" }} allow="fullscreen; autoplay; encrypted-media; picture-in-picture" />
                    ) : (
                      <div className="nf-player" style={sfondo(corso, 1)}>
                        <span style={{ width: 74, height: 74, borderRadius: "50%", background: "rgba(255,255,255,.2)", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 30 }}>▶</span>
                        <b style={{ fontSize: 18 }}>{lezione.titolo}</b>
                        <span style={{ color: "rgba(255,255,255,.75)", fontSize: 13 }}>Qui il video della lezione{esempio ? " (esempio)" : ""} · {lezione.durata}</span>
                      </div>
                    )
                  ) : (
                    <div className="nf-player" style={sfondo(corso, 1)}>
                      <span style={{ fontSize: 40 }}>🔒</span>
                      <b style={{ fontSize: 18 }}>{lezione.titolo}</b>
                      <span style={{ color: "rgba(255,255,255,.8)", fontSize: 13, maxWidth: 460, textAlign: "center" }}>Fa parte di «{corso.titolo}». Sbloccalo per guardare tutte le lezioni.</span>
                      <button className="nf-sblocca" onClick={sblocca}>Sblocca · {corso.prezzo || "premium"}</button>
                    </div>
                  )}
                </div>
              ) : (
                corso.copertina ? (
                <div style={{ position: "relative" }}>
                  <img src={corso.copertina} alt={corso.titolo} style={{ width: "100%", aspectRatio: "16 / 9", objectFit: "cover", display: "block" }} />
                  <button className="nf-chiudi" aria-label="Chiudi" onClick={() => setAcCorso("")}>×</button>
                  <div className="riga" style={{ gap: 10, flexWrap: "wrap", position: "absolute", left: 24, bottom: 20 }}>
                    {primaLibera(corso) && <button className="nf-play" onClick={() => setAcLezione({ corso: corso.id, lezione: primaLibera(corso).id })}>▶ Guarda</button>}
                    {!corso.gratuito && !corso.sbloccato && <button className="nf-sblocca" onClick={sblocca}>{corso.prezzo ? `Sblocca · ${corso.prezzo}` : "Accedi al corso"}</button>}
                  </div>
                </div>
                ) : (
                <div className="nf-mtesta">
                  <div className="sfondo" style={sfondo(corso, 0)} />
                  <button className="nf-chiudi" aria-label="Chiudi" onClick={() => setAcCorso("")}>×</button>
                  <div>
                    <img src={LOGO_ACADEMY} alt="Accademia dell'Advertising" style={{ width: 54, height: 54, marginBottom: 10, filter: "drop-shadow(0 4px 12px rgba(0,0,0,.4))" }} />
                    <h2>{corso.titolo}</h2>
                    <div className="riga" style={{ gap: 10, flexWrap: "wrap" }}>
                      {primaLibera(corso) && <button className="nf-play" onClick={() => setAcLezione({ corso: corso.id, lezione: primaLibera(corso).id })}>▶ {corso.gratuito || corso.sbloccato ? "Guarda" : "Guarda l'anteprima"}</button>}
                      {!corso.gratuito && !corso.sbloccato && <button className="nf-sblocca" onClick={sblocca}>{corso.prezzo ? `Sblocca · ${corso.prezzo}` : "Accedi al corso"}</button>}
                    </div>
                  </div>
                </div>
                )
              )}
              <div className="nf-corpo">
                {corso.copertina && !lezione && <h2 style={{ margin: "14px 0 0", fontSize: 24, letterSpacing: "-.02em" }}>{corso.titolo}</h2>}
                <div className="nf-dati">
                  {corso.esempio ? <span className="tag">esempio</span> : corso.gratuito ? <span className="tag" data-k="attivo">gratis</span> : corso.sbloccato ? <span className="tag" data-k="attivo">sbloccato</span> : corso.prezzo ? <span className="tag" data-k="alta">🔒 {corso.prezzo}</span> : null}
                  <span>{(ACADEMY_MACRO.find((m2) => m2.v === (corso.macro || "corsi")) || {}).l}</span><span>·</span><span>{corso.livello}</span>
                  {nLez(corso) > 0 && <><span>·</span><span>{nLez(corso)} lezioni</span><span>·</span><span>seguite {nViste(corso)} di {nLez(corso)}</span></>}
                  {lezione && lezioneLibera(corso, lezione) && (
                    <label className="riga" style={{ gap: 6, marginLeft: "auto" }}><input type="checkbox" style={{ width: 16, height: 16, accentColor: "var(--acc)" }} checked={!!viste[lezione.id]} onChange={(e) => segna(lezione.id, e.target.checked)} /> lezione seguita</label>
                  )}
                </div>
                <div style={{ color: "var(--ink2)", lineHeight: 1.55 }}>{corso.descrizione}</div>
                <div>
                  <h3 style={{ margin: "6px 0 8px", fontSize: 18 }}>Lezioni</h3>
                  {!nLez(corso) && (
                    <div className="meta-msg" style={{ borderRadius: 8, display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                      <span>Le lezioni di questo corso compariranno qui quando colleghiamo la piattaforma dell'Accademia.</span>
                      <span className="sp" />
                      {urlOk(ac.url) && <a className="btn ghost mini" href={ac.url} target="_blank" rel="noopener noreferrer" style={{ textDecoration: "none" }}>Apri sulla piattaforma ↗</a>}
                    </div>
                  )}
                  {(corso.lezioni || []).map((l, k) => {
                    const libera = lezioneLibera(corso, l);
                    return (
                      <button key={l.id} className="nf-ep" onClick={() => setAcLezione({ corso: corso.id, lezione: l.id })} style={acLezione && acLezione.lezione === l.id ? { background: "var(--surf2)" } : undefined}>
                        <span className="n">{k + 1}</span>
                        <span className="mini" style={sfondo(corso, k)}>{viste[l.id] ? "✓" : libera ? "▶" : "🔒"}</span>
                        <span><b>{l.titolo}</b><span className="sub">{l.gratuita && !corso.gratuito ? "Anteprima gratuita" : libera ? "Disponibile" : "Da sbloccare"}</span></span>
                        <span className="sub mono dur">{l.durata}</span>
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  };

  const SezioneWinner = () => {
    const bm = benchmarkDi(attivo);
    const cambiaBm = (campi) => setData((d) => ({ ...d, clienti: d.clienti.map((z) => (z.id === attivo.id ? { ...z, winner: { ...benchmarkDi(z), ...campi } } : z)) }));
    const foto = fotografie(winDati.ads);
    const scelta = foto.find((f) => `${f.data}|${f.fonte}` === winFoto) || foto[0];
    const diFoto = (f) => (f ? winDati.ads.filter((r) => r.data === f.data && r.fonte === f.fonte) : []);
    const precedente = scelta ? foto.find((f) => f.fonte === scelta.fonte && f.data < scelta.data) : null;
    /* Le copie della stessa creatività (anche in pausa) si sommano; una classifica per pubblico */
    const base = (f) => (bm.raggruppa ? raggruppaCopie(diFoto(f)) : diFoto(f));
    const PUBBLICI = [["freddo", "Traffico freddo"], ["retargeting", "Retargeting"], ["altro", "Altro"], ["", "Pubblico non indicato"]];
    const perPubblico = (l, pub) => l.filter((a) => (a.pubblico || "") === pub);
    const oraBase = base(scelta), primaBase = precedente ? base(precedente) : [];
    const sezioni = PUBBLICI.filter(([pub]) => (!bm.pubblico || bm.pubblico === pub) && perPubblico(oraBase, pub).length).map(([pub, titolo]) => {
      const { righe, esclusi } = classifica(perPubblico(oraBase, pub), { ...bm, pubblico: "" });
      return { pub, titolo, esclusi, lista: conAndamento(righe, classifica(perPubblico(primaBase, pub), { ...bm, pubblico: "" }).righe) };
    });
    const lista = sezioni.flatMap((x) => x.lista);
    const esclusi = sezioni.reduce((x, y) => x + y.esclusi, 0);
    const tot = lista.reduce((x, a) => ({ sp: x.sp + a.spesa, cli: x.cli + a.clienti, app: x.app + a.appuntamenti }), { sp: 0, cli: 0, app: 0 });
    const colore = (r) => (r <= 1 ? "pos" : r > 1.3 ? "neg" : "");
    const campo = (k, etichetta, passo = 1) => (
      <label style={{ display: "grid", gap: 4 }}>
        <span className="eyebrow">{etichetta}</span>
        <input type="number" min="0" step={passo} value={bm[k]} style={{ width: 110 }}
          onChange={(e) => cambiaBm({ [k]: Math.max(0, Number(e.target.value) || 0) })} />
      </label>
    );
    const nomeCliAd = (id) => (winIds.length > 1 ? nomeCli(id) : "");
    return (
      <>
        {SuggerimentoAcademy({ sez: "winner" })}
        <section className="card">
          <header>
            <h2>Benchmark</h2>
            <span className="sp eyebrow">punteggio = media dei tre costi divisi per il loro riferimento · sotto 1 batte i riferimenti</span>
            <button className="btn ghost mini" onClick={() => cambiaBm({ ...BENCH_BASE })}>Valori predefiniti</button>
          </header>
          <div className="riga" style={{ gap: 14, flexWrap: "wrap", padding: "14px 18px", alignItems: "flex-end" }}>
            {campo("cpl", "Costo per lead")}
            {campo("app", "Costo per appuntamento")}
            {campo("cliente", "Costo per cliente")}
            {campo("minClienti", "Minimo clienti")}
            {campo("minSpesa", "Minimo spesa", 50)}
            <label style={{ display: "grid", gap: 4 }}>
              <span className="eyebrow">Ordina per</span>
              <select value={bm.ordine} onChange={(e) => cambiaBm({ ordine: e.target.value })} style={{ width: "auto" }}>
                {ORDINI.map((o) => <option key={o.v} value={o.v}>{o.l}</option>)}
              </select>
            </label>
            <label style={{ display: "grid", gap: 4 }}>
              <span className="eyebrow">Pubblico</span>
              <select value={bm.pubblico} onChange={(e) => cambiaBm({ pubblico: e.target.value })} style={{ width: "auto" }}>
                <option value="">tutti</option><option value="freddo">freddo</option><option value="retargeting">retargeting</option><option value="altro">altro</option>
              </select>
            </label>
            <label className="riga" style={{ gap: 6, paddingBottom: 8 }} title="Somma i numeri di tutte le copie della stessa creatività, anche quelle in pausa">
              <input type="checkbox" style={{ width: 16, height: 16, accentColor: "var(--acc)" }} checked={!!bm.raggruppa} onChange={(e) => cambiaBm({ raggruppa: e.target.checked })} /> raggruppa le copie
            </label>
            <label className="riga" style={{ gap: 6, paddingBottom: 8 }}>
              <input type="checkbox" style={{ width: 16, height: 16, accentColor: "var(--acc)" }} checked={!!bm.soloAttive} onChange={(e) => cambiaBm({ soloAttive: e.target.checked })} /> {bm.raggruppa ? "solo con una copia accesa" : "solo attive"}
            </label>
          </div>
        </section>

        {winDati.errore && <div className="meta-msg" data-tipo="errore">{winDati.errore}</div>}
        {!foto.length ? (
          <section className="card">
            <div className="empty">
              <b>{winDati.carico ? "Carico…" : "Ancora nessun dato sugli annunci"}</b>
              {!winDati.carico && <>I numeri arrivano dal Claude del cliente attraverso il connettore del CRM: crealo in Impostazioni, riquadro «Connettore per Claude», aggiungilo in claude.ai e chiedi al Claude del cliente di salvare le winner ads nel CRM. La Guida, capitolo «Winner ads», ha tutti i passaggi.</>}
            </div>
          </section>
        ) : (
          <>
            <div className="kpis">
              <div className="kpi"><IcoKpi t="Annunci" /><div className="eyebrow">In classifica</div><div className="v">{int(lista.length)}</div><div className="s">{esclusi ? `${int(esclusi)} fuori dai filtri` : "tutti dentro i filtri"}</div></div>
              <div className="kpi"><IcoKpi t="Spesa" /><div className="eyebrow">Spesa</div><div className="v">{eur0(tot.sp)}</div><div className="s">degli annunci in classifica</div></div>
              <div className="kpi"><IcoKpi t="Clienti chiusi" /><div className="eyebrow">Clienti</div><div className="v">{int(tot.cli)}</div><div className="s">costo per cliente {tot.cli ? eur(tot.sp / tot.cli) : "—"}</div></div>
              <div className="kpi"><IcoKpi t="Tasso" /><div className="eyebrow">Chiusura</div><div className="v">{tot.app ? pctv((tot.cli / tot.app) * 100, 0) : "—"}</div><div className="s">clienti su appuntamenti</div></div>
            </div>
            <section className="card">
              <header>
                <h2>Fotografia</h2>
                <span className="sp eyebrow">{scelta.fonte} · {scelta.periodo || "periodo non indicato"}{precedente ? ` · confronto con il ${dataIt(precedente.data)}` : ""}{bm.raggruppa ? " · copie raggruppate" : ""}</span>
                <select value={`${scelta.data}|${scelta.fonte}`} onChange={(e) => setWinFoto(e.target.value)} aria-label="Fotografia" style={{ width: "auto" }}>
                  {foto.map((f) => <option key={`${f.data}|${f.fonte}`} value={`${f.data}|${f.fonte}`}>{dataIt(f.data)} · {f.fonte} · {f.n} annunci</option>)}
                </select>
                <button className="btn ghost mini" disabled={winDati.carico} onClick={() => setWinGiro((x) => x + 1)}>{winDati.carico ? "Carico…" : "Ricarica"}</button>
              </header>
            </section>
            {sezioni.map((sz) => (
              <section className="card" key={sz.pub || "nessuno"}>
                <header>
                  <h2>Winner ads · {sz.titolo}</h2>
                  <span className="sp eyebrow">{int(sz.lista.length)} in classifica · {(ORDINI.find((o) => o.v === bm.ordine) || ORDINI[0]).l.toLowerCase()}{sz.esclusi ? ` · ${int(sz.esclusi)} fuori dai filtri` : ""}</span>
                </header>
                <div className="wrap"><table className="tbl">
                  <thead><tr>
                    <th className="r">#</th><th>{bm.raggruppa ? "Creatività" : "Annuncio"}</th>
                    {[["spesa", "Spesa"], ["lead", "Lead"], ["cpl", "CPL"], ["appuntamenti", "Appunt."], ["cpa", "Costo/app."], ["clienti", "Clienti"], ["cpc", "Costo/cliente"], ["chiusura", "Chiusura"], ["punteggio", "Punteggio"]].map(([v, l]) => (
                      <th key={v} className="r clic" onClick={() => cambiaBm({ ordine: v })} aria-sort={bm.ordine === v ? (ORDINI.find((o) => o.v === v).su ? "ascending" : "descending") : undefined}
                        title={`Ordina per ${(ORDINI.find((o) => o.v === v) || {}).l || l}`} style={bm.ordine === v ? { color: "var(--ink)" } : undefined}>
                        {l}{bm.ordine === v ? (ORDINI.find((o) => o.v === v).su ? " ↑" : " ↓") : ""}
                      </th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {sz.lista.map((a) => (
                      <tr key={a.cliente_id + a.ad_id}>
                        <td className="r mono">{a.pos}{a.salita ? <div className={"sub " + (a.salita > 0 ? "pos" : "neg")}>{a.salita > 0 ? `▲${a.salita}` : `▼${-a.salita}`}</div> : a.salita === null && precedente ? <div className="sub">nuovo</div> : null}</td>
                        <td>
                          <b>{a.nome}</b>
                          <div className="sub">{[
                            a.copie > 1 ? `${a.copie} copie · ${a.copieAttive ? `${a.copieAttive} ${a.copieAttive === 1 ? "accesa" : "accese"}` : "tutte in pausa"}` : a.attiva ? "" : "in pausa",
                            a.campagna, nomeCliAd(a.cliente_id),
                          ].filter(Boolean).join(" · ")}</div>
                          {a.note && <div className="sub" style={{ maxWidth: 420 }}>{a.note}</div>}
                        </td>
                        <td className="r">{eur0(a.spesa)}</td>
                        <td className="r">{int(a.lead)}</td>
                        <td className={"r " + colore(a.rCpl)}>{a.cpl ? eur(a.cpl) : "—"}</td>
                        <td className="r">{int(a.appuntamenti)}</td>
                        <td className={"r " + colore(a.rApp)}>{a.cpa ? eur(a.cpa) : "—"}</td>
                        <td className="r">{int(a.clienti)}</td>
                        <td className={"r " + colore(a.rCli)}>{a.cpc ? eur(a.cpc) : "—"}</td>
                        <td className="r">{a.chiusura ? pctv(a.chiusura * 100, 0) : "—"}</td>
                        <td className={"r mono " + colore(a.punteggio)}><b>{fmt(a.punteggio)}</b></td>
                      </tr>
                    ))}
                  </tbody>
                </table></div>
                {!sz.lista.length && <div className="empty"><b>Nessuna creatività dentro i filtri</b>Abbassa il minimo di clienti o di spesa, o togli «solo con una copia accesa».</div>}
              </section>
            ))}
          </>
        )}

        <section className="card">
          <header><h2>Analisi</h2><span className="sp eyebrow">salvate dal Claude del cliente · le più recenti in alto</span></header>
          {!winDati.testi.length && <div className="empty"><b>Nessuna analisi</b>Quando il Claude del cliente usa «salva_analisi», il testo compare qui con la data.</div>}
          {winDati.testi.map((x, i) => (
            <details key={x.id} open={i === 0} style={{ borderTop: "1px solid var(--line)", padding: "12px 18px" }}>
              <summary style={{ cursor: "pointer" }}>
                <b>{x.titolo}</b> <span className="sub">· {dataIt(String(x.creato).slice(0, 10))} {String(x.creato).slice(11, 16)}{winIds.length > 1 ? ` · ${nomeCli(x.cliente_id)}` : ""}{x.tipo && x.tipo !== "altro" ? ` · ${x.tipo.replace("_", " ")}` : ""}</span>
              </summary>
              <div style={{ marginTop: 10, lineHeight: 1.55, fontSize: 13.5 }}>{testoAnalisi(x.testo)}</div>
              <div style={{ marginTop: 10 }}>
                <button className="btn danger mini" onClick={async () => {
                  if (!window.confirm(`Eliminare l'analisi «${x.titolo}»?`)) return;
                  try { await analisi.eliminaTesto(x.id); setWinGiro((g) => g + 1); } catch (e) { setWinDati((w) => ({ ...w, errore: e.message || String(e) })); }
                }}>Elimina</button>
              </div>
            </details>
          ))}
        </section>
      </>
    );
  };

  const SezioneAdv = () => {
    const base = (attivo ? S.campagne : m.campagne).filter((k) =>
      (attivo || !fAdv.cli || k.clienteId === fAdv.cli) &&
      (!fAdv.account || (k.metaAccount || "__mano") === fAdv.account) &&
      (!fAdv.piattaforma || k.piattaforma === fAdv.piattaforma) &&
      (!fAdv.obiettivo || k.obiettivo === fAdv.obiettivo) &&
      (!fAdv.stato || k.stato === fAdv.stato) &&
      (!fAdv.funnel || (fAdv.funnel === "__senza" ? !k.funnelId : k.funnelId === fAdv.funnel)) &&
      (!fAdv.linea || (fAdv.linea === "__senza" ? !k.linea : k.linea === fAdv.linea)) &&
      (!fAdv.faseCorrente || (fAdv.faseCorrente === "__valutazione"
        ? !!inValutazione(k) : faseCampagna(k).fase === fAdv.faseCorrente)) &&
      (!fAdv.segnale || semaforo(data, k).segnale === fAdv.segnale) &&
      cerca(k.nome + " " + nomeCli(k.clienteId))
    ).sort((a, b) => {
      /* L'ordine predefinito è dalla più recente: è quella su cui devi decidere oggi */
      const dataA = a.inizio || "", dataB = b.inizio || "";
      switch (ordAdv) {
        case "vecchie": return dataA.localeCompare(dataB) || a.nome.localeCompare(b.nome, "it");
        case "spesaGiu": return b.d.sp - a.d.sp;
        case "spesaSu": return a.d.sp - b.d.sp;
        case "cplSu": return (a.d.cpl || Infinity) - (b.d.cpl || Infinity);
        case "roasGiu": return (b.d.roas || 0) - (a.d.roas || 0);
        case "leadGiu": return b.d.ld - a.d.ld;
        case "nome": return a.nome.localeCompare(b.nome, "it");
        default: return dataB.localeCompare(dataA) || a.nome.localeCompare(b.nome, "it");
      }
    });
    const somma = (l) => l.reduce((s, k) => ({
      sp: s.sp + k.d.sp, imp: s.imp + k.d.imp, cl: s.cl + k.d.cl, clk: s.clk + k.d.clk, ld: s.ld + k.d.ld,
      cv: s.cv + k.d.cv, rc: s.rc + k.d.rc,
      bg: s.bg + (k.budget || 0), fq: s.fq + (k.d.freq || 0) * k.d.imp, n: s.n + 1,
    }), { sp: 0, imp: 0, cl: 0, clk: 0, ld: 0, cv: 0, rc: 0, bg: 0, fq: 0, n: 0 });
    /* Le medie non si mediano: si ricalcolano dai totali, altrimenti la campagna
       da cento euro pesa quanto quella da diecimila. */
    const derivate = (g) => ({
      ...g, ctr: g.imp ? (g.cl / g.imp) * 100 : 0, cpc: g.cl ? g.sp / g.cl : 0, cpm: g.imp ? (g.sp / g.imp) * 1000 : 0,
      ctrLink: g.imp ? ((g.clk || g.cl) / g.imp) * 100 : 0,
      cpcLink: (g.clk || g.cl) ? g.sp / (g.clk || g.cl) : 0,
      cpl: g.ld ? g.sp / g.ld : 0,
      cpa: g.cv ? g.sp / g.cv : 0, roas: g.sp ? g.rc / g.sp : 0, freq: g.imp ? g.fq / g.imp : 0,
    });
    const T = derivate(somma(base));

    const nomeFunnel = (id) => (data.funnel.find((f) => f.id === id) || {}).nome || "";
    /* Contratto e incassato arrivano dalle schede chiusura collegate alla campagna,
       filtrate per data quando si guarda una singola finestra. */
    const daChiusure = (idCamp, dal, al) => {
      const l = data.chiusure.filter((c) => c.campagnaId === idCamp && (!dal || (c.data >= dal && c.data <= al)));
      return {
        n: l.length,
        contratto: l.reduce((x, c) => x + parseNumIntl(c.valoreContratto), 0),
        incassato: l.reduce((x, c) => x + parseNumIntl(c.incassato), 0),
      };
    };
    /* Se contratto e incassato sono scritti nella scheda valgono quelli,
       altrimenti li ricavo dalle schede chiusura collegate. */
    const soldiDi = (rec, idCamp, dal, al) => {
      const ct = parseNumIntl(rec && rec.valoreContratto);
      const ic = parseNumIntl(rec && rec.incassato);
      if (ct + ic > 0) return { n: 0, contratto: ct, incassato: ic, manuale: true };
      return daChiusure(idCamp, dal, al);
    };
    const soldiTot = (l) => l.reduce((x, k) => {
      const q = soldiDi(k, k.id);
      return { n: x.n + q.n, contratto: x.contratto + q.contratto, incassato: x.incassato + q.incassato };
    }, { n: 0, contratto: 0, incassato: 0 });

    const linee = [...new Set((attivo ? S.campagne : m.campagne).map((k) => k.linea).filter(Boolean))].sort();
    const perLinea = (() => {
      const mappa = new Map();
      base.forEach((k) => {
        const key = k.linea || "__senza";
        mappa.set(key, [...(mappa.get(key) || []), k]);
      });
      return [...mappa.entries()]
        .map(([id, l]) => ({ id, nome: id === "__senza" ? "Senza linea" : id, ...derivate(somma(l)), ...soldiTot(l) }))
        .sort((a, b) => b.sp - a.sp);
    })();
    const mescolate = perLinea.filter((x) => x.id !== "__senza").length > 1;

    const perFunnel = (() => {
      const mappa = new Map();
      base.forEach((k) => {
        const key = k.funnelId || "__senza";
        mappa.set(key, [...(mappa.get(key) || []), k]);
      });
      return [...mappa.entries()]
        .map(([id, l]) => ({ id, nome: id === "__senza" ? "Senza funnel" : nomeFunnel(id) || "Funnel eliminato", ...derivate(somma(l)), ...soldiTot(l) }))
        .sort((a, b) => b.sp - a.sp);
    })();

    const perPiat = new Map();
    base.forEach((k) => {
      const key = et(PIATTAFORME, k.piattaforma);
      const g = perPiat.get(key) || { sp: 0, imp: 0, cl: 0, clk: 0, ld: 0, cv: 0, rc: 0, n: 0, al: 0, contratto: 0, incassato: 0 };
      const q = soldiDi(k, k.id);
      g.sp += k.d.sp; g.imp += k.d.imp; g.cl += k.d.cl; g.clk += k.d.clk; g.ld += k.d.ld; g.cv += k.d.cv; g.rc += k.d.rc; g.n += 1; g.al += veri(k.al).length;
      g.contratto += q.contratto; g.incassato += q.incassato;
      perPiat.set(key, g);
    });
    const piattaforme = [...perPiat.entries()].sort((a, b) => b[1].sp - a[1].sp);

    /* Numeri della fase scelta: se la campagna non ha quella finestra, non ha numeri da mostrare */
    const numeriFase = (k) => {
      if (fase === "totale") return { d: k.d, quando: "", dal: "", al: "" };
      const fs = (data.periodi || []).filter((pz) => pz.campagnaId === k.id).sort((a, b) => (a.inizio < b.inizio ? -1 : 1));
      let scelti = [];
      if (fase === "test" || fase === "matura") scelti = fs.filter((pz) => pz.tipo === fase);
      else if (fase === "settimane") scelti = fs.filter((pz) => pz.tipo === "attiva");
      else if (fase === "ultima") { const w = fs.filter((pz) => pz.tipo === "attiva"); scelti = w.length ? [w[w.length - 1]] : []; }
      if (!scelti.length) return null;
      const g = derivate(somma(scelti.map((pz) => ({ ...pz, d: calc(pz), budget: 0 }))));
      const sd = scelti.reduce((x, pz) => {
        const q = soldiDi(pz, k.id, pz.inizio, pz.fine);
        return { n: x.n + q.n, contratto: x.contratto + q.contratto, incassato: x.incassato + q.incassato };
      }, { n: 0, contratto: 0, incassato: 0 });
      return { d: g, soldi: sd, dal: scelti[0].inizio, al: scelti[scelti.length - 1].fine,
        quando: `${dataIt(scelti[0].inizio)}–${dataIt(scelti[scelti.length - 1].fine)}${scelti.length > 1 ? ` · ${scelti.length} finestre` : ""}` };
    };
    const conFase = base.map((k) => ({ k, f: numeriFase(k) })).filter((x) => x.f !== null);
    const TF = derivate(somma(conFase.map((x) => ({ ...x.k, d: x.f.d }))));

    const inCorso = base.filter((k) => k.stato === "attiva" && k.budgetGiorno > 0);
    const cliMarg = attivo || (fAdv.cli ? cliente(fAdv.cli) : null);
    const prof = cliMarg ? profitto(T.rc, T.sp, cliMarg) : null;

    const salvaRilevazione = () => {
      const g = oggi();
      const nuove = base.filter((k) => ["attiva", "test"].includes(k.stato)).map((k) => ({
        id: uid(), data: g, clienteId: k.clienteId, campagnaId: k.id,
        spesa: k.spesa || 0, impression: k.impression || 0, click: k.click || 0,
        conversioni: k.conversioni || 0, ricavi: k.ricavi || 0, frequenza: k.frequenza || 0,
        lead: k.lead || 0, clickLink: k.clickLink || 0,
      }));
      if (!nuove.length) return;
      upd("rilevazioni", (l) => [...l.filter((r) => !(r.data === g && nuove.some((n) => n.campagnaId === r.campagnaId))), ...nuove]);
    };
    /* Anche le voci non ancora collegate a una campagna (es. i quattro annunci Core 4),
       altrimenti non compaiono da nessuna parte e non si possono aprire per collegarle. */
    const creativi = (attivo ? S.creativi : data.creativi)
      .filter((c) => base.some((k) => k.id === c.campagnaId) || (!c.campagnaId && (!attivo || ambito.includes(c.clienteId) || c.clienteId === attivo.id)))
      .map((c) => {
        const d = calc(c);
        const contratti = data.chiusure
          .filter((w) => w.clienteId === c.clienteId && (w.adset === c.nome || w.annuncio === c.nome))
          .reduce((s, w) => s + (w.valoreContratto || 0), 0);
        /* Il tetto del CPL è quello della campagna madre: ad set e annunci
           non hanno benchmark propri, rispondono a quello della campagna. */
        const madre = base.find((k) => k.id === c.campagnaId) || {};
        return { ...c, d, contratti, tetto: tettoCpl(madre.bmk) };
      })
      .sort((a, b) => b.d.sp - a.d.sp);
    const nomeCampagna = (id) => (data.campagne.find((k) => k.id === id) || {}).nome || "—";

    /* Fuori dai clienti solo i totali per cliente: campagne, ad set e finestre si guardano dentro il cliente */
    if (!attivo) {
      const tutte = m.campagne.filter((k) => cerca(k.nome + " " + nomeCli(k.clienteId)));
      /* Contratti e incassato: dalle campagne (schede collegate o importi scritti) se ci sono,
         altrimenti dal Ciclo di vendita del cliente. Il ritorno si calcola sui soli mesi che hanno
         contratti, con la spesa di quegli stessi mesi: spesa da sempre contro contratti di un anno
         darebbe un numero che non corrisponde a niente. */
      /* Una riga per cliente principale con i totali suoi e dei sottoclienti; la freccia apre i sottoclienti */
      const calcola = (z, liv, ids, extra = {}) => {
        const l = tutte.filter((k) => ids.includes(k.clienteId));
        if (!l.length) return null;
        const camp = soldiTot(l);
        const mv = righeMensili(ids).filter((r) => r.contratto || r.incassato);
        const ctV = mv.reduce((x, r) => x + r.contratto, 0), icV = mv.reduce((x, r) => x + r.incassato, 0), spV = mv.reduce((x, r) => x + r.spesa, 0);
        const daCiclo = !(camp.contratto + camp.incassato) && ctV + icV > 0;
        return { z, liv, key: z.id + (extra.solo ? ":solo" : ""), ...extra, ...derivate(somma(l)), n: l.length, accese: l.filter((k) => ["attiva", "test"].includes(k.stato)).length,
          contratto: daCiclo ? ctV : camp.contratto, incassato: daCiclo ? icV : camp.incassato, daCiclo,
          ritorno: daCiclo && spV ? ctV / spV : 0, mesiV: daCiclo ? mv.length : 0 };
      };
      const righeCli = ordinaGerarchia().filter((x) => x.liv === 0).flatMap(({ z }) => {
        const sotto = ordinaGerarchia(z.id, 1);
        const r = calcola(z, 0, [z.id, ...sotto.map((x) => x.z.id)], { nSotto: sotto.length });
        if (!r) return [];
        if (!sotto.length || !apertiAdv[z.id]) return [r];
        const propria = calcola(z, 1, [z.id], { solo: true });
        return [r, ...(propria ? [propria] : []), ...sotto.map((x) => calcola(x.z, x.liv, [x.z.id])).filter(Boolean)];
      });
      const TT = derivate(somma(tutte));
      const ST = righeCli.filter((r) => r.liv === 0).reduce((x, r) => ({ contratto: x.contratto + r.contratto, incassato: x.incassato + r.incassato }), { contratto: 0, incassato: 0 });
      const entra = (id) => { cambiaAccount(id); setVista("campagne"); };
      return conValuta(valutaComune(tutte.map((k) => k.clienteId)), () => (
        <>
          <div className="kpis">
            <div className="kpi"><IcoKpi t="Spesa del periodo" /><div className="eyebrow">Spesa</div><div className="v">{eur0(TT.sp)}</div><div className="s">{righeCli.filter((r) => r.liv === 0).length} clienti principali · {tutte.length} campagne</div></div>
            <div className="kpi"><IcoKpi /><div className="eyebrow">Lead</div><div className="v">{int(TT.ld)}</div><div className="s">CPL medio {TT.ld ? eur(TT.cpl) : "—"}</div></div>
            <div className="kpi"><IcoKpi /><div className="eyebrow">Conversioni</div><div className="v">{int(TT.cv)}</div><div className="s">CPA medio {TT.cv ? eur(TT.cpa) : "—"}</div></div>
            <div className="kpi"><IcoKpi /><div className="eyebrow">Incassato</div><div className="v">{eur0(ST.incassato)}</div><div className="s">contratti {eur0(ST.contratto)}{TT.sp && TT.rc ? ` · ROAS ${fmt(TT.roas)}` : ""}</div></div>
          </div>
          {SuggerimentoAcademy({ sez: "advertising" })}
          <section className="card">
            <header><h2>Advertising per cliente</h2><span className="sp eyebrow">clienti principali con i loro sottoclienti · ▸ apre i sottoclienti</span></header>
            <div className="wrap"><table className="tbl">
              <thead><tr><th>Cliente</th><th className="r">Campagne</th><th className="r">Spesa</th><th className="r">Lead</th><th className="r">CPL</th><th className="r">Conversioni</th><th className="r" title="ROAS dichiarato dalla piattaforma; se manca, contratti su spesa dei mesi del Ciclo di vendita">ROAS</th><th className="r">Contratti</th><th className="r">Incassato</th></tr></thead>
              <tbody>
                {mapValuta(righeCli, (r) => r.z.id, (r) => (
                  <tr key={r.key} className="clic" onClick={() => entra(r.z.id)} title="Apri l'advertising del cliente">
                    <td style={{ paddingLeft: 14 + r.liv * 18 }}>
                      {r.nSotto > 0 && (
                        <button className="adv-freccia" aria-expanded={!!apertiAdv[r.z.id]} aria-label={apertiAdv[r.z.id] ? "Chiudi i sottoclienti" : "Apri i sottoclienti"}
                          onClick={(e) => { e.stopPropagation(); setApertiAdv((x) => ({ ...x, [r.z.id]: !x[r.z.id] })); }}>{apertiAdv[r.z.id] ? "▾" : "▸"}</button>
                      )}
                      {r.liv ? <span className="sub">↳ </span> : null}<b>{r.z.nome}</b>
                      {r.solo && <div className="sub">solo l'account principale</div>}
                      {r.nSotto > 0 && <div className="sub">con {int(r.nSotto)} {r.nSotto === 1 ? "sottocliente" : "sottoclienti"}</div>}
                    </td>
                    <td className="r">{int(r.n)}<div className="sub mono">{int(r.accese)} accese</div></td>
                    <td className="r">{eur0(r.sp)}</td>
                    <td className="r">{int(r.ld)}</td>
                    <td className="r">{r.ld ? eur(r.cpl) : "—"}</td>
                    <td className="r">{int(r.cv)}</td>
                    <td className="r">{r.sp && r.rc ? fmt(r.roas) : r.ritorno ? <>{fmt(r.ritorno)}×<div className="sub">su contratti</div></> : "—"}</td>
                    <td className="r">{eur0(r.contratto)}{r.daCiclo ? <div className="sub">ciclo di vendita · {int(r.mesiV)} {r.mesiV === 1 ? "mese" : "mesi"}</div> : null}</td>
                    <td className="r">{eur0(r.incassato)}</td>
                  </tr>
                ))}
              </tbody>
            </table></div>
            {!righeCli.length && <div className="empty"><b>Nessuna campagna</b>Le campagne si aggiungono dentro il cliente, o arrivano da Meta.</div>}
          </section>
        </>
      ));
    }

    /* Più account pubblicitari per lo stesso cliente: ogni campagna sa da quale account arriva
       (metaAccount); quelle scritte a mano stanno insieme sotto "Inserite a mano". */
    const nomeAccountAdv = (id) => {
      if (id === "__mano") return "Inserite a mano";
      const a = metaAccount.find((x) => String(x.account_id) === String(id));
      return a ? a.nome || `act_${id}` : `act_${id}`;
    };
    const accountAdv = attivo ? [...new Set((attivo ? S.campagne : []).map((k) => k.metaAccount || "__mano"))]
      .map((id) => ({ id, nome: nomeAccountAdv(id) })).sort((a, b) => a.nome.localeCompare(b.nome, "it")) : [];
    /* Totali per obiettivo, con sotto il dettaglio di ogni account: le situazioni restano distinte,
       ma dove l'obiettivo è lo stesso i numeri si sommano e le medie si ricalcolano sui totali */
    const perObiettivoAccount = (() => {
      if (!attivo || accountAdv.length < 2) return [];
      const ob = new Map();
      base.forEach((k) => { const o = k.obiettivo || "altro"; ob.set(o, [...(ob.get(o) || []), k]); });
      return [...ob.entries()].map(([o, l]) => {
        const acc = new Map();
        l.forEach((k) => { const a = k.metaAccount || "__mano"; acc.set(a, [...(acc.get(a) || []), k]); });
        return {
          o, nome: et(OBIETTIVI, o) || "Altro", tot: { ...derivate(somma(l)), ...soldiTot(l) }, n: l.length,
          account: [...acc.entries()].map(([a, la]) => ({ id: a, nome: nomeAccountAdv(a), n: la.length, accese: la.filter((k) => ["attiva", "test", "errore"].includes(k.stato)).length,
            bloccato: (la.find((k) => k.accountBloccato) || {}).accountBloccato || "", ...derivate(somma(la)), ...soldiTot(la) })).sort((x, y) => y.sp - x.sp),
        };
      }).sort((a, b) => b.tot.sp - a.tot.sp);
    })();

    /* Dentro un cliente con sottoclienti: prima le campagne del cliente, poi una riga chiusa per
       ogni sottocliente (con i suoi totali) che si apre con la freccia e mostra le sue campagne */
    const listaAdv = (righe) => {
      if (!attivo) return righe;
      const sotto = figli(attivo.id);
      const proprie = righe.filter((x) => !sotto.some((z) => [z.id, ...discendenti(z.id)].includes(x.k.clienteId)));
      const gruppi = sotto.map((z) => {
        const ids = [z.id, ...discendenti(z.id)];
        const items = righe.filter((x) => ids.includes(x.k.clienteId));
        return items.length ? { gruppo: z, items, tot: { ...derivate(somma(items.map((x) => ({ d: x.f.d, budget: x.k.budget })))), ...soldiTot(items.map((x) => x.k)) } } : null;
      }).filter(Boolean).sort((a, b) => b.tot.sp - a.tot.sp);
      if (!gruppi.length) return righe;
      return [...proprie, ...gruppi.flatMap((g) => [g, ...(apertiAdv[g.gruppo.id] ? g.items : [])])];
    };
    const rigaGruppoAdv = ({ gruppo: z, items, tot: g }) => {
      const aperto = !!apertiAdv[z.id];
      const accese = items.filter((x) => ["attiva", "test"].includes(x.k.stato)).length;
      return (
        <tr key={"g" + z.id} className="clic adv-gruppo" onClick={() => setApertiAdv((x) => ({ ...x, [z.id]: !x[z.id] }))}
          title={aperto ? "Chiudi le campagne del sottocliente" : "Apri le campagne del sottocliente"}>
          <td style={{ paddingRight: 0, width: 22 }}><span className="adv-freccia" aria-expanded={aperto}>{aperto ? "▾" : "▸"}</span></td>
          <td colSpan={3}><b>{z.nome}</b><div className="sub">sottocliente · {int(items.length)} campagne, {int(accese)} accese</div></td>
          <td className="r">{eur0(g.sp)}</td>
          <td className="r">{g.imp ? pctv(g.ctrLink) : "—"}</td>
          <td className="r">{(g.clk || g.cl) ? eur(g.cpcLink) : "—"}</td>
          <td className="r">{int(g.ld)}</td>
          <td className="r">{g.ld ? eur(g.cpl) : "—"}</td>
          <td className="r">{int(g.cv)}</td>
          <td className="r">{g.cv ? eur(g.cpa) : "—"}</td>
          <td className="r">{g.rc ? fmt(g.roas) : "—"}</td>
          <td className="r">{g.contratto ? eur0(g.contratto) : "—"}</td>
          <td className="r">{g.contratto ? eur0(g.incassato) : "—"}</td>
          <td /><td /><td />
        </tr>
      );
    };

    return conValuta(valutaComune(base.map((k) => k.clienteId)), () => (
      <>
        <div className="kpis">
          <div className="kpi"><IcoKpi t="Spesa del periodo" /><div className="eyebrow">Spesa del periodo</div><div className="v">{eur0(T.sp)}</div><div className="s">{base.length} campagne · budget {eur0(T.bg)}{mescolate && !fAdv.linea ? " · più linee insieme" : ""}</div></div>
          <div className="kpi"><IcoKpi t="CTR link" /><div className="eyebrow">CTR link</div><div className="v">{T.imp ? pctv(T.ctrLink) : "—"}</div><div className="s">{int(T.clk || T.cl)} clic su {int(T.imp)} impression</div></div>
          <div className="kpi">
            <IcoKpi />
            <div className="eyebrow">{T.ld ? "CPL medio" : "CPA medio"}</div>
            <div className="v">{T.ld ? eur(T.cpl) : T.cv ? eur(T.cpa) : "—"}</div>
            <div className="s">{T.ld ? `${int(T.ld)} lead` : `${int(T.cv)} conversioni`} · CPC link {eur(T.cpcLink)}</div>
          </div>
          <div className="kpi">
            <IcoKpi />
            <div className="eyebrow">{cliMarg && margineNetto(cliMarg) ? "Margine stimato" : "ROAS"}</div>
            <div className="v" style={{ color: prof !== null && prof < 0 ? "var(--acc)" : "var(--ink)" }}>
              {cliMarg && margineNetto(cliMarg) ? eur0(prof) : T.rc ? `${fmt(T.sp ? T.rc / T.sp : 0)}×` : "—"}
            </div>
            <div className="s">
              {soldiTot(base).contratto
                ? `contratto ${eur0(soldiTot(base).contratto)} · incassato ${eur0(soldiTot(base).incassato)}`
                : cliMarg && margineNetto(cliMarg)
                  ? `ROAS ${fmt(T.sp ? T.rc / T.sp : 0)}× · pareggio a ${fmt(roasPareggio(cliMarg))}×`
                  : T.rc ? `${eur0(T.rc)} di ricavi tracciati` : "campagne senza ricavi"}
            </div>
          </div>
        </div>

        <section className="card">
          <header>
            <div className="ctrl">
              {(() => {
                const attiviF = [fAdv.cli, fAdv.piattaforma, fAdv.obiettivo, fAdv.stato,
                  fAdv.funnel, fAdv.linea, fAdv.faseCorrente, fAdv.segnale].filter(Boolean).length;
                const etOrd = {
                  recenti: "dalla più recente", vecchie: "dalla più vecchia", spesaGiu: "spesa ↓", spesaSu: "spesa ↑",
                  leadGiu: "lead ↓", cplSu: "CPL ↑", roasGiu: "ROAS ↓", nome: "nome A-Z",
                }[ordAdv];
                const etVista = { totale: "", test: "solo Test", matura: "solo Matura", ultima: "ultima settimana", settimane: "settimane attive" }[fase];
                return (
                  <>
                    <button className="btn ghost mini" data-on={filtriAperti ? "1" : "0"}
                      onClick={() => setFiltriAperti(!filtriAperti)} aria-expanded={filtriAperti}>
                      {filtriAperti ? "▾" : "▸"} Filtri{attiviF ? ` · ${attiviF}` : ""}
                    </button>
                    <span className="sub" style={{ fontSize: 12 }}>
                      {base.length} campagne · {etOrd}{etVista ? ` · ${etVista}` : ""}
                      {fAdv.segnale && (
                        <span className="riga" style={{ display: "inline-flex", gap: 6, marginLeft: 8, verticalAlign: "middle" }}>
                          <span style={{ width: 8, height: 8, borderRadius: "50%", display: "inline-block",
                            background: (SEGNALI.find((g) => g.v === fAdv.segnale) || {}).c }} />
                          {(SEGNALI.find((g) => g.v === fAdv.segnale) || {}).l}
                        </span>
                      )}
                    </span>
                    {(attiviF > 0 || ordAdv !== "recenti" || fase !== "totale") && (
                      <button className="btn ghost mini" onClick={() => {
                        setFAdv({ cli: "", piattaforma: "", obiettivo: "", stato: "", funnel: "", linea: "", faseCorrente: "", segnale: "" });
                        setOrdAdv("recenti"); setFase("totale");
                      }}>Azzera</button>
                    )}
                  </>
                );
              })()}
            </div>
            <span className="sp" />
            {attivo && (
              <button className="btn" onClick={() => setRichiesta({ nome: attivo.nome, testo: componiRichiesta(attivo) })}>
                Chiedi i dati aggiornati
              </button>
            )}
            <button className="btn ghost mini" onClick={() => {
              const ids = base.map((k) => k.id);
              chiedi({
                titolo: "Applicare le soglie del framework?",
                testo: `Imposto su ${ids.length} campagne: CTR minimo 0,80% · frequenza massima 4,50 · CPM massimo al 140% della media dell'account${T.imp ? ` (${eur((T.sp / T.imp) * 1000 * 1.4)})` : ""}.\n\nROAS minimo e CPA massimo non li tocco: quelli dipendono dal margine del cliente e li decidi tu.`,
                etichetta: "Applica",
                azione: () => upd("campagne", (l) => l.map((k) => (ids.includes(k.id)
                  ? { ...k, bmk: { ...k.bmk, ctrMin: 0.8, freqMax: 4.5, cpmMax: T.imp ? Math.round((T.sp / T.imp) * 1000 * 1.4 * 100) / 100 : k.bmk.cpmMax } }
                  : k))),
              });
            }}>Soglie del framework</button>
            {/* Meta riporta i lead come "risultati": importati, finiscono anche fra le
                conversioni e il CPA diventa una copia del CPL. Questo li separa. */}
            {base.some((k) => k.d.cv > 0 && k.d.cv === k.d.ld) && (
              <button className="btn ghost mini" onClick={() => {
                const ids = base.filter((k) => k.d.cv > 0 && k.d.cv === k.d.ld).map((k) => k.id);
                chiedi({
                  titolo: "Separare conversioni e lead?",
                  testo: `Su ${ids.length} ${ids.length === 1 ? "campagna" : "campagne"} il numero delle conversioni è identico a quello dei lead, e il CPA risulta uguale al CPL. Quasi sempre è l'importazione: Meta chiama "risultati" i lead, e finiscono in tutte e due le colonne.\n\nAzzero le conversioni, nelle campagne e nelle loro finestre. I lead restano, il CPL resta, il CPA torna vuoto finché non scrivi tu quante di quelle persone sono diventate appuntamenti o clienti.`,
                  etichetta: "Azzera le conversioni",
                  azione: () => {
                    upd("campagne", (l) => l.map((k) => (ids.includes(k.id) ? { ...k, conversioni: 0, esitoManuale: "si" } : k)));
                    upd("periodi", (l) => l.map((pz) => (ids.includes(pz.campagnaId) ? { ...pz, conversioni: 0 } : pz)));
                  },
                });
              }}>Separa conversioni dai lead</button>
            )}
            <button className="btn ghost mini" onClick={salvaRilevazione} title="Fotografa i numeri di oggi per poterli confrontare domani">
              Salva rilevazione
            </button>
          </header>
          {filtriAperti && (
            <div className="pannello-filtri">
              {attivo && accountAdv.length > 1 && (
                <div className="f">
                  <label htmlFor="fl-acc">Account pubblicitario</label>
                  <select id="fl-acc" value={fAdv.account || ""} onChange={(e) => setFAdv({ ...fAdv, account: e.target.value })}>
                    <option value="">Tutti gli account</option>
                    {accountAdv.map((a) => <option key={a.id} value={a.id}>{a.nome}</option>)}
                  </select>
                </div>
              )}
              {!attivo && (
                <div className="f">
                  <label htmlFor="fl-cli">Cliente</label>
                  <select id="fl-cli" value={fAdv.cli} onChange={(e) => setFAdv({ ...fAdv, cli: e.target.value })}>
                    <option value="">Tutti i clienti</option>
                    {attivi.map((z) => <option key={z.id} value={z.id}>{z.nome}</option>)}
                  </select>
                </div>
              )}
              <div className="f">
                <label htmlFor="fl-ord">Ordina per</label>
                <select id="fl-ord" value={ordAdv} onChange={(e) => setOrdAdv(e.target.value)}>
                  <option value="recenti">Dalla più recente</option>
                  <option value="vecchie">Dalla più vecchia</option>
                  <option value="spesaGiu">Spesa, dalla più alta</option>
                  <option value="spesaSu">Spesa, dalla più bassa</option>
                  <option value="leadGiu">Lead, dai più numerosi</option>
                  <option value="cplSu">CPL, dal più basso</option>
                  <option value="roasGiu">ROAS, dal più alto</option>
                  <option value="nome">Nome, A-Z</option>
                </select>
              </div>
              <div className="f">
                <label htmlFor="fl-fasec">Fase corrente</label>
                <select id="fl-fasec" value={fAdv.faseCorrente} onChange={(e) => setFAdv({ ...fAdv, faseCorrente: e.target.value })}>
                  <option value="">Tutte le fasi</option>
                  <option value="test">In fase Test</option>
                  <option value="matura">In fase Matura</option>
                  <option value="attiva">In fase Attiva</option>
                  <option value="chiusa">Chiuse</option>
                  <option value="errore">Ferme (pagamento o account)</option>
                  <option value="__valutazione">In valutazione dopo modifica</option>
                  <option value="nd">Senza data di inizio</option>
                </select>
              </div>
              <div className="f">
                <label htmlFor="fl-segn">Semaforo</label>
                <select id="fl-segn" value={fAdv.segnale} onChange={(e) => setFAdv({ ...fAdv, segnale: e.target.value })}>
                  <option value="">Tutti i segnali</option>
                  {SEGNALI.map((g) => <option key={g.v} value={g.v}>{g.l}</option>)}
                </select>
              </div>
              <div className="f">
                <label htmlFor="fl-stato">Stato</label>
                <select id="fl-stato" value={fAdv.stato} onChange={(e) => setFAdv({ ...fAdv, stato: e.target.value })}>
                  <option value="">Tutti gli stati</option>
                  {STATI_CAMPAGNA.map((s) => <option key={s.v} value={s.v}>{s.l}</option>)}
                </select>
              </div>
              <div className="f">
                <label htmlFor="fl-vista">Numeri da mostrare</label>
                <select id="fl-vista" value={fase} onChange={(e) => setFase(e.target.value)}>
                  <option value="totale">Totale campagna</option>
                  <option value="test">Solo fase Test</option>
                  <option value="matura">Solo fase Matura</option>
                  <option value="ultima">Ultima settimana attiva</option>
                  <option value="settimane">Tutte le settimane attive</option>
                </select>
                <span className="hint">cambia i numeri, non quali campagne vedi</span>
              </div>
              <div className="f">
                <label htmlFor="fl-piatt">Piattaforma</label>
                <select id="fl-piatt" value={fAdv.piattaforma} onChange={(e) => setFAdv({ ...fAdv, piattaforma: e.target.value })}>
                  <option value="">Tutte le piattaforme</option>
                  {PIATTAFORME.map((p) => <option key={p.v} value={p.v}>{p.l}</option>)}
                </select>
              </div>
              <div className="f">
                <label htmlFor="fl-ob">Obiettivo</label>
                <select id="fl-ob" value={fAdv.obiettivo} onChange={(e) => setFAdv({ ...fAdv, obiettivo: e.target.value })}>
                  <option value="">Tutti gli obiettivi</option>
                  {OBIETTIVI.map((o) => <option key={o.v} value={o.v}>{o.l}</option>)}
                </select>
              </div>
              <div className="f">
                <label htmlFor="fl-funnel">Funnel</label>
                <select id="fl-funnel" value={fAdv.funnel} onChange={(e) => setFAdv({ ...fAdv, funnel: e.target.value })}>
                  <option value="">Tutti i funnel</option>
                  {(attivo ? S.funnel : data.funnel).map((f) => <option key={f.id} value={f.id}>{f.nome}</option>)}
                  <option value="__senza">Senza funnel</option>
                </select>
              </div>
              {linee.length > 0 && (
                <div className="f">
                  <label htmlFor="fl-linea">Linea di business</label>
                  <select id="fl-linea" value={fAdv.linea} onChange={(e) => setFAdv({ ...fAdv, linea: e.target.value })}>
                    <option value="">Tutte le linee</option>
                    {linee.map((l) => <option key={l} value={l}>{l}</option>)}
                    <option value="__senza">Senza linea</option>
                  </select>
                </div>
              )}
            </div>
          )}
          <div className="totali">
            {(() => {
              const V = fase === "totale" ? T : TF;
              const q = soldiTot(base);
              const voci = [
                ["Campagne", fase === "totale" ? int(base.length) : int(conFase.length)],
                ["Speso", eur(V.sp)],
                ["CTR link", pctv(V.ctrLink)],
                ["CPC link", eur(V.cpcLink)],
                ...(V.ld ? [["Lead", int(V.ld)], ["CPL", eur(V.cpl)]] : []),
                ["Conversioni", int(V.cv)],
                ...(V.cv ? [["CPA", eur(V.cpa)]] : []),
                ...(V.rc ? [["ROAS", `${fmt(V.roas)}×`]] : []),
                ...(q.contratto ? [["Contratto", eur0(q.contratto)], ["Incassato", eur0(q.incassato)]] : []),
                ["Frequenza", fmt(V.freq)],
              ];
              return voci.map(([l, v]) => (
                <div className="tot" key={l}><span className="eyebrow">{l}</span><b>{v}</b></div>
              ));
            })()}
          </div>
          <div className="wrap"><table className="tbl">
            <thead><tr>
              <th /><th>Campagna</th><th>Funnel</th><th>Fase</th><th className="r">Speso</th><th className="r">CTR link</th><th className="r">CPC link</th><th className="r">Lead</th><th className="r">CPL</th>
              <th className="r">Conv.</th><th className="r">CPA</th><th className="r">ROAS</th><th className="r">Contratto</th><th className="r">Incassato</th><th className="r">Andamento</th><th className="r">Avvisi</th><th />
            </tr></thead>
            <tbody>{mapValuta(listaAdv(fase === "totale" ? base.map((k) => ({ k, f: { d: k.d, quando: "" } })) : conFase), (x) => (x.gruppo ? x.gruppo.id : x.k.clienteId), (x) => x.gruppo ? rigaGruppoAdv(x) : (({ k, f }) => {
              const iv = intervallo(k, ultimaRilevazione(data, k));
              const D = f.d;
              return (
                <tr key={k.id} className="clic" onClick={() => apri("campagne", data.campagne.find((x) => x.id === k.id))}>
                  <td style={{ paddingRight: 0, width: 22 }}>
                    {(() => {
                      const sg = semaforo(data, k);
                      return <span title={`${sg.etichetta} — ${sg.perche}`}
                        style={{ width: 9, height: 9, borderRadius: "50%", background: sg.colore, display: "inline-block" }} />;
                    })()}
                  </td>
                  <td>
                    <div className="riga" style={{ gap: 7 }}><b>{k.nome}</b><Campanella k={k} /></div>
                    <div className="sub">
                      {k.linea ? <span className="pill" style={{ marginRight: 6 }}>{k.linea}</span> : null}
                      {k.inizioMeta && (
                        <span className="pill" style={{ marginRight: 6 }}
                          title={`In scheda risulta ${dataItA(k.inizioMeta)}, ma ha già speso dal ${dataItA(k.inizio)}: i calcoli partono da lì.`}>
                          dal {dataIt(k.inizio)}
                        </span>
                      )}
                      {et(PIATTAFORME, k.piattaforma)} · {attivo ? et(OBIETTIVI, k.obiettivo) : nomeCli(k.clienteId)}
                      {attivo && accountAdv.length > 1 ? ` · ${nomeAccountAdv(k.metaAccount || "__mano")}` : ""}
                    </div>
                  </td>
                  <td className="sub">{k.funnelId ? (nomeFunnel(k.funnelId) || "—") : <span className="sub">—</span>}</td>
                  <td>
                    <span className="tag" data-k={k.stato}>{etStato(k)}</span>
                    {k.accountBloccato ? <div className="acc-bloccato" title={k.stato === "errore" ? "La campagna era accesa: l'account l'ha fermata" : "La campagna era già spenta: l'account non la ferma, ma se la riaccendi non erogherà"}>⚠ account {k.accountBloccato}{k.stato !== "errore" ? " · campagna già spenta" : ""}</div> : null}
                    <div className="sub">{fase === "totale" ? faseCampagna(k).l : f.quando}</div>
                    {(() => {
                      const pr = valutazionePronta(k);
                      if (pr) return (
                        <div className="sub" style={{ color: "#C99A3A" }}
                          title={`${pr.nota ? `${pr.nota} · ` : ""}i tre giorni sono scaduti il ${dataItA(pr.scaduta)}`}>
                          valutazione pronta{pr.da > 0 ? ` da ${pr.da} ${pr.da === 1 ? "giorno" : "giorni"}` : " da oggi"}
                        </div>
                      );
                      const w = inValutazione(k);
                      return w ? (
                        <div className="sub" style={{ color: "#C99A3A" }}
                          title={`${w.nota ? `${w.nota} · ` : ""}modificata il ${dataItA(w.dal)}, non toccare fino al ${dataItA(w.fino)}`}>
                          in valutazione dopo modifica · {w.restano === 1 ? "ancora un giorno" : `ancora ${w.restano} giorni`}
                        </div>
                      ) : null;
                    })()}
                  </td>
                  <td className="r">{eur(D.sp)}<div className="sub mono">
                      {k.budgetGiorno ? `${eur(budgetDelGiorno(k, oggi()))}/g` : `di ${eur(k.budget)}`}
                      {(k.storicoBudget || []).length > 1 ? ` · ${(k.storicoBudget || []).length} cambi` : ""}
                    </div></td>
                  <td className={"r " + (k.bmk.ctrMin && (D.ctrLink || D.ctr) < k.bmk.ctrMin ? "neg" : "")}>
                    {pctv(D.ctrLink || D.ctr)}
                    {D.soloTotali && <div className="sub">clic totali</div>}
                  </td>
                  <td className="r">{eur(D.cpcLink || D.cpc)}</td>
                  <td className="r">{D.ld ? int(D.ld) : "—"}</td>
                  <td className={"r " + (tettoCpl(k.bmk) && D.ld && D.cpl > tettoCpl(k.bmk) ? "neg" : "")}>{D.ld ? eur(D.cpl) : "—"}</td>
                  <td className="r">
                    {int(D.cv)}
                    {k.esitoManuale === "si" && <div className="sub">a mano</div>}
                    {D.cv > 0 && D.cv === D.ld && k.esitoManuale !== "si" && <div className="sub">= lead</div>}
                  </td>
                  <td className={"r " + (k.bmk.cpaMax && D.cv && D.cpa > k.bmk.cpaMax ? "neg" : "")}>{D.cv ? eur(D.cpa) : "—"}</td>
                  <td className={"r " + (D.rc ? (k.bmk.roasMin && D.roas < k.bmk.roasMin ? "neg" : "pos") : "")}>{D.rc ? `${fmt(D.roas)}×` : "—"}</td>
                  {(() => {
                    const q = fase === "totale" ? soldiDi(k, k.id) : (f.soldi || soldiDi(null, k.id, f.dal, f.al));
                    return (
                      <>
                        <td className="r">{q.contratto ? eur0(q.contratto) : "—"}{q.n > 0 && <div className="sub">{q.n} {q.n === 1 ? "chiusura" : "chiusure"}</div>}</td>
                        <td className={"r " + (q.contratto && q.incassato < q.contratto ? "neg" : "")}>{q.contratto ? eur0(q.incassato) : "—"}</td>
                      </>
                    );
                  })()}
                  <td className="r">
                    {(() => {
                      const w = faseCampagna(k).fase === "attiva" ? finestraAnalisi(data, k) : null;
                      if (!w) return null;
                      return <div className="sub mono">finestra {w.giorni} g · {eur(w.sp)}</div>;
                    })()}
                    {!iv ? <span className="sub">—</span>
                      : iv.cv > 0 && D.cv > 0
                        ? <span className={iv.cpa > D.cpa ? "neg" : "pos"}>{iv.cpa > D.cpa ? "▲" : "▼"} {eur(iv.cpa)}</span>
                        : iv.rc > 0
                          ? <span className={iv.roas < D.roas ? "neg" : "pos"}>{fmt(iv.roas)}×</span>
                          : <span className="sub">{eur(iv.sp)}</span>}
                    {iv && <div className="sub mono">dal {dataIt(iv.da)}</div>}
                  </td>
                  <td className="r">{veri(k.al).length ? <span className="tag" data-k={k.al.some((a) => a.gravita === "critica") ? "late" : ""}>{veri(k.al).length}</span> : <span className="sub">ok</span>}</td>
                  <td className="r">
                    <button className="btn ghost mini" onClick={(e) => { e.stopPropagation(); setCampFasi(campFasi === k.id ? "" : k.id); }}>Fasi</button>
                    <button className="btn ghost mini" title="Toglila da medie e totali"
                      onClick={(e) => { e.stopPropagation(); upd("campagne", (l) => l.map((x) => (x.id === k.id ? { ...x, archiviata: "si" } : x))); }}>Archivia</button>
                  </td>
                </tr>
              );
            })(x))}</tbody>
            {base.length > 0 && (
              <tfoot>
                <tr className="totale">
                  <td />
                  <td><b>Tutte insieme</b><div className="sub">{fase === "totale" ? `${base.length} campagne come se fossero una` : `${conFase.length} campagne · ${et([{ v: "test", l: "fase Test" }, { v: "matura", l: "fase Matura" }, { v: "ultima", l: "ultima settimana" }, { v: "settimane", l: "settimane attive" }], fase)}`}</div></td>
                  <td className="sub">{perFunnel.length} funnel</td>
                  <td />
                  <td className="r"><b>{eur((fase === "totale" ? T : TF).sp)}</b><div className="sub mono">di {eur(T.bg)}</div></td>
                  <td className="r"><b>{pctv((fase === "totale" ? T : TF).ctrLink)}</b></td>
                  <td className="r"><b>{eur((fase === "totale" ? T : TF).cpcLink)}</b></td>
                  <td className="r"><b>{(fase === "totale" ? T : TF).ld ? int((fase === "totale" ? T : TF).ld) : "—"}</b></td>
                  <td className="r"><b>{(fase === "totale" ? T : TF).ld ? eur((fase === "totale" ? T : TF).cpl) : "—"}</b></td>
                  <td className="r"><b>{int((fase === "totale" ? T : TF).cv)}</b></td>
                  <td className="r"><b>{(fase === "totale" ? T : TF).cv ? eur((fase === "totale" ? T : TF).cpa) : "—"}</b></td>
                  <td className="r"><b>{(fase === "totale" ? T : TF).rc ? `${fmt((fase === "totale" ? T : TF).roas)}×` : "—"}</b></td>
                  <td className="r"><b>{(() => { const q = fase === "totale" ? soldiTot(base) : conFase.reduce((x, y) => ({ contratto: x.contratto + (y.f.soldi ? y.f.soldi.contratto : 0), incassato: x.incassato + (y.f.soldi ? y.f.soldi.incassato : 0) }), { contratto: 0, incassato: 0 }); return q.contratto ? eur0(q.contratto) : "—"; })()}</b></td>
                  <td className="r"><b>{(() => { const q = fase === "totale" ? soldiTot(base) : conFase.reduce((x, y) => ({ contratto: x.contratto + (y.f.soldi ? y.f.soldi.contratto : 0), incassato: x.incassato + (y.f.soldi ? y.f.soldi.incassato : 0) }), { contratto: 0, incassato: 0 }); return q.contratto ? eur0(q.incassato) : "—"; })()}</b></td>
                  <td className="r sub">freq. {fmt(T.freq)}</td>
                  <td className="r sub">{veri(base.flatMap((k) => k.al)).length}</td>
                  <td />
                </tr>
              </tfoot>
            )}
          </table></div>
          {!base.length && <div className="empty"><b>Nessuna campagna</b>{attivo ? "Aggiungine una con il pulsante in alto." : "Con questi filtri non compare nulla."}</div>}
        </section>

        <div className="two">
          <section className="card">
            <header><h2>Ritmo di erogazione</h2><span className="sp eyebrow">spesa reale al giorno contro budget impostato</span></header>
            <div className="port">
              {mapValuta(inCorso, (k) => k.clienteId, (k) => {
                const E = erogazioneAttesa(data, k);
                const bgOggi = budgetDelGiorno(k, oggi());
                const attesoGiorno = E.giorni && E.atteso > 0 ? E.atteso / E.giorni : bgOggi;
                const spesaGiorno = E.giorni ? k.d.sp / E.giorni : 0;
                const speso = attesoGiorno > 0 ? spesaGiorno / attesoGiorno : 0;
                if (E.sospetto) return (
                  <div className="pr clic" key={k.id} onClick={() => apri("campagne", data.campagne.find((x) => x.id === k.id))}>
                    <div className="who"><div><b>{k.nome}</b><span className="sub neg">date da controllare</span></div></div>
                    <div className="track" />
                    <div className="val sub">{int(E.giorni)} giorni dal {dataItA(E.dal)}</div>
                  </div>
                );
                const cambi = (k.storicoBudget || []).length;
                const ultimoCambio = cambi > 1 ? (k.storicoBudget || []).slice(-1)[0] : null;
                return (
                  <div className="pr clic" key={k.id} onClick={() => apri("campagne", data.campagne.find((x) => x.id === k.id))}>
                    <div className="who"><div><b>{k.nome}</b><span className="sub">
                      {eur(budgetDelGiorno(k, oggi()))}/g{ultimoCambio ? ` · cambiato il ${dataIt(ultimoCambio.data)}` : ""}
                    </span></div></div>
                    <div className="track">
                      <div className="fill" data-sotto={speso > 1.2 || speso < 0.75 ? "1" : "0"} style={{ width: `${Math.min(100, (speso / 1.5) * 100)}%` }} />
                      <div className="mark" style={{ left: "66.6%" }} />
                    </div>
                    <div className="val">{eur(spesaGiorno)}<div className="sub" style={{ fontFamily: "var(--ui)" }}>
                      al giorno · {fmt(speso * 100, 0)}% di {eur(attesoGiorno)} · {E.giorni} g{E.daFasi ? " dalle fasi" : ""}
                    </div></div>
                  </div>
                );
              })}
              {!inCorso.length && <div className="empty"><b>Nessuna campagna attiva con budget</b>Il ritmo si calcola su budget e date.</div>}
            </div>
          </section>

          <section className="card">
            <header><h2>Riepilogo per piattaforma</h2></header>
            <div className="wrap">
              <table className="tbl">
                <thead><tr><th>Piattaforma</th><th className="r">Spesa</th><th className="r">CTR link</th><th className="r">Lead</th><th className="r">CPL</th><th className="r">CPA</th><th className="r">ROAS</th><th className="r">Contratto</th><th className="r">Incassato</th></tr></thead>
                <tbody>
                  {piattaforme.map(([nome, g]) => (
                    <tr key={nome}>
                      <td><b>{nome}</b><div className="sub">{g.n} campagne{g.al ? ` · ${g.al} avvisi` : ""}</div></td>
                      <td className="r">{eur(g.sp)}</td>
                      <td className="r">{g.imp ? pctv(((g.clk || g.cl) / g.imp) * 100) : "—"}</td>
                      <td className="r">{g.ld ? int(g.ld) : "—"}</td>
                      <td className="r">{g.ld ? eur(g.sp / g.ld) : "—"}</td>
                      <td className="r">{g.cv ? eur(g.sp / g.cv) : "—"}</td>
                      <td className="r">{g.rc ? `${fmt(g.rc / g.sp)}×` : "—"}</td>
                      <td className="r">{g.contratto ? eur0(g.contratto) : "—"}</td>
                      <td className="r">{g.contratto ? eur0(g.incassato) : "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {!piattaforme.length && <div className="empty"><b>Nessun dato</b>Serve almeno una campagna.</div>}
          </section>
        </div>

        {perObiettivoAccount.length > 0 && (
          <section className="card">
            <header><h2>Per obiettivo e account</h2><span className="sp eyebrow">{accountAdv.length} account · totali sommati solo a parità di obiettivo</span></header>
            <div className="wrap"><table className="tbl">
              <thead><tr><th>Obiettivo · account</th><th className="r">Campagne</th><th className="r">Spesa</th><th className="r">CTR link</th><th className="r">Lead</th><th className="r">CPL</th><th className="r">Conv.</th><th className="r">CPA</th><th className="r">ROAS</th><th className="r">Contratto</th></tr></thead>
              <tbody>
                {perObiettivoAccount.flatMap((g) => [
                  <tr key={g.o} className="adv-gruppo">
                    <td><b>{g.nome}</b><div className="sub">totale di {g.account.length} {g.account.length === 1 ? "account" : "account"}</div></td>
                    <td className="r">{int(g.n)}</td>
                    <td className="r"><b>{eur0(g.tot.sp)}</b></td>
                    <td className="r">{g.tot.imp ? pctv(g.tot.ctrLink) : "—"}</td>
                    <td className="r">{g.tot.ld ? int(g.tot.ld) : "—"}</td>
                    <td className="r"><b>{g.tot.ld ? eur(g.tot.cpl) : "—"}</b></td>
                    <td className="r">{g.tot.cv ? int(g.tot.cv) : "—"}</td>
                    <td className="r"><b>{g.tot.cv ? eur(g.tot.cpa) : "—"}</b></td>
                    <td className="r">{g.tot.rc ? `${fmt(g.tot.roas)}×` : "—"}</td>
                    <td className="r">{g.tot.contratto ? eur0(g.tot.contratto) : "—"}</td>
                  </tr>,
                  ...g.account.map((a) => (
                    <tr key={g.o + a.id} className="clic" onClick={() => { setFAdv({ ...fAdv, account: a.id, obiettivo: g.o }); setFiltriAperti(true); }}
                      title="Mostra solo le campagne di questo account con questo obiettivo">
                      <td style={{ paddingLeft: 32 }}><span className="sub">↳ </span>{a.nome}
                        <div className="sub">{int(a.accese)} accese{a.bloccato ? <span className="neg"> · account {a.bloccato}</span> : null}</div></td>
                      <td className="r">{int(a.n)}</td>
                      <td className="r">{eur0(a.sp)}</td>
                      <td className="r">{a.imp ? pctv(a.ctrLink) : "—"}</td>
                      <td className="r">{a.ld ? int(a.ld) : "—"}</td>
                      <td className="r">{a.ld ? eur(a.cpl) : "—"}</td>
                      <td className="r">{a.cv ? int(a.cv) : "—"}</td>
                      <td className="r">{a.cv ? eur(a.cpa) : "—"}</td>
                      <td className="r">{a.rc ? `${fmt(a.roas)}×` : "—"}</td>
                      <td className="r">{a.contratto ? eur0(a.contratto) : "—"}</td>
                    </tr>
                  )),
                ])}
              </tbody>
            </table></div>
            <div className="nota">Un clic su un account filtra l'elenco delle campagne su quell'account e quell'obiettivo. Obiettivi diversi non si sommano: un CPL di lead e un CPA di vendite non sono la stessa misura.</div>
          </section>
        )}

        {campFasi && (() => {
          const k = base.find((x) => x.id === campFasi) || m.campagne.find((x) => x.id === campFasi);
          if (!k) return null;
          const fasi = (data.periodi || []).filter((pz) => pz.campagnaId === k.id)
            .sort((a, b) => (a.inizio < b.inizio ? -1 : 1));
          const etichetta = (pz, i, tutte) => {
            if (pz.tipo !== "attiva") return et(TIPI_PERIODO, pz.tipo);
            const n = tutte.filter((x) => x.tipo === "attiva" && x.inizio <= pz.inizio).length;
            return `Attiva · settimana ${n}`;
          };
          const tot = derivate(somma(fasi.map((pz) => ({ ...pz, d: calc(pz), budget: 0 }))));

          const genera = () => {
            if (!k.inizio) return;
            const esistenti = fasi.map((pz) => pz.inizio);
            const nuovi = [];
            const agg = (tipo, dal, al) => {
              if (esistenti.includes(dal) || dal > oggi()) return;
              nuovi.push({ id: uid(), clienteId: k.clienteId, campagnaId: k.id, tipo, inizio: dal, fine: al > oggi() ? oggi() : al,
                spesa: 0, impression: 0, click: 0, clickLink: 0, lead: 0, conversioni: 0, ricavi: 0, valoreContratto: 0, incassato: 0, frequenza: 0, note: "" });
            };
            agg("test", k.inizio, sommaGiorni(k.inizio, TEST_1 - 1));
            agg("matura", sommaGiorni(k.inizio, TEST_1), sommaGiorni(k.inizio, TEST_2 - 1));
            let dal = sommaGiorni(k.inizio, TEST_2);
            let guardia = 0;
            while (dal <= oggi() && guardia < 60) { agg("attiva", dal, sommaGiorni(dal, FINESTRA - 1)); dal = sommaGiorni(dal, FINESTRA); guardia++; }
            if (!nuovi.length) return;
            chiedi({
              titolo: "Creare le finestre mancanti?",
              testo: `Preparo ${nuovi.length} periodi vuoti seguendo le date: ${TEST_1} giorni di Test, ${TEST_2 - TEST_1} di Matura, poi settimane Attive da ${FINESTRA} giorni fino a oggi. I numeri li inserisci tu.`,
              etichetta: "Crea",
              azione: () => upd("periodi", (l) => [...l, ...nuovi]),
            });
          };

          return conValuta(valutaCliente(data, k.clienteId), () => (
            <section className="card">
              <header>
                <h2>Fasi di {k.nome}</h2>
                <span className="sp eyebrow">{k.inizio ? `partita il ${dataIt(k.inizio)}` : "senza data di inizio"}</span>
                {k.inizio && <button className="btn ghost mini" onClick={genera}>Genera le finestre</button>}
                <button className="btn ghost mini" onClick={() => { setCampFasi(k.id); apri("periodi", { campagnaId: k.id, clienteId: k.clienteId }, true); }}>+ Fase</button>
                <button className="btn ghost mini" onClick={() => setCampFasi("")}>Chiudi</button>
              </header>
              <div className="wrap">
                <table className="tbl">
                  <thead><tr>
                    <th>Fase</th><th>Periodo</th><th className="r">Speso</th><th className="r">Clic link</th><th className="r">CTR link</th>
                    <th className="r">CPC link</th><th className="r">Lead</th><th className="r">CPL</th><th className="r">Conv.</th><th className="r">CPA</th><th className="r">ROAS</th><th className="r">Contratto</th><th className="r">Incassato</th>
                  </tr></thead>
                  <tbody>
                    {fasi.map((pz, i) => {
                      const q = calc(pz);
                      return (
                        <tr key={pz.id} className="clic" onClick={() => apri("periodi", pz)}>
                          <td>
                            <b>{etichetta(pz, i, fasi)}</b>
                            {pz.note && <div className="sub">{pz.note}</div>}
                          </td>
                          <td className="sub">{dataIt(pz.inizio)}–{dataIt(pz.fine)}</td>
                          <td className="r">{eur(q.sp)}</td>
                          <td className="r">{int(q.clk || q.cl)}</td>
                          <td className={"r " + (k.bmk.ctrMin && q.imp > 1000 && (q.ctrLink || q.ctr) < k.bmk.ctrMin ? "neg" : "")}>{pctv(q.ctrLink || q.ctr)}</td>
                          <td className="r">{eur(q.cpcLink || q.cpc)}</td>
                          <td className="r">{q.ld ? int(q.ld) : "—"}</td>
                          <td className={"r " + (tettoCpl(k.bmk) && q.ld && q.cpl > tettoCpl(k.bmk) ? "neg" : "")}>{q.ld ? eur(q.cpl) : "—"}</td>
                          <td className="r">{int(q.cv)}</td>
                          <td className={"r " + (k.bmk.cpaMax && q.cv && q.cpa > k.bmk.cpaMax ? "neg" : "")}>{q.cv ? eur(q.cpa) : "—"}</td>
                          <td className={"r " + (q.rc ? (k.bmk.roasMin && q.roas < k.bmk.roasMin ? "neg" : "pos") : "")}>{q.rc ? `${fmt(q.roas)}×` : "—"}</td>
                          <td className="r">{soldiDi(pz, k.id, pz.inizio, pz.fine).contratto ? eur0(soldiDi(pz, k.id, pz.inizio, pz.fine).contratto) : "—"}</td>
                          <td className="r">{soldiDi(pz, k.id, pz.inizio, pz.fine).contratto ? eur0(soldiDi(pz, k.id, pz.inizio, pz.fine).incassato) : "—"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                  {fasi.length > 0 && (
                    <tfoot><tr className="totale">
                      <td><b>Totale campagna</b><div className="sub">{fasi.length} finestre</div></td>
                      <td className="sub">{dataIt(fasi[0].inizio)}–{dataIt(fasi[fasi.length - 1].fine)}</td>
                      <td className="r"><b>{eur(tot.sp)}</b></td>
                      <td className="r"><b>{int(tot.clk || tot.cl)}</b></td>
                      <td className="r"><b>{pctv(tot.ctrLink)}</b></td>
                      <td className="r"><b>{eur(tot.cpcLink)}</b></td>
                      <td className="r"><b>{tot.ld ? int(tot.ld) : "—"}</b></td>
                      <td className="r"><b>{tot.ld ? eur(tot.cpl) : "—"}</b></td>
                      <td className="r"><b>{int(tot.cv)}</b></td>
                      <td className="r"><b>{tot.cv ? eur(tot.cpa) : "—"}</b></td>
                      <td className="r"><b>{tot.rc ? `${fmt(tot.roas)}×` : "—"}</b></td>
                      <td className="r"><b>{soldiDi(k, k.id).contratto ? eur0(soldiDi(k, k.id).contratto) : "—"}</b></td>
                      <td className="r"><b>{soldiDi(k, k.id).contratto ? eur0(soldiDi(k, k.id).incassato) : "—"}</b></td>
                    </tr></tfoot>
                  )}
                </table>
              </div>
              {!fasi.length && (
                <div className="empty">
                  <b>Nessuna finestra registrata</b>
                  Con "Genera le finestre" preparo Test, Matura e settimane Attive secondo la data di partenza: tu inserisci solo i numeri.
                </div>
              )}
              {fasi.length > 0 && (
                <div className="nota">
                  Quando una campagna ha le fasi registrate, i suoi totali in tabella non sono più quelli scritti a mano nella scheda:
                  sono la somma di queste finestre, con le medie ricalcolate.
                </div>
              )}
            </section>
          ));
        })()}

        {(() => {
          const arch = m.archiviate.filter((k) => (attivo ? ambito.includes(k.clienteId) : (!fAdv.cli || k.clienteId === fAdv.cli)));
          const chiudibili = base.filter((k) => k.stato === "conclusa");
          const depAdv = (data.deposito || []).filter((v2) => v2.motivo === "advertising"
            && (attivo ? (v2.ids || []).some((x) => ambito.includes(x)) : true));
          if (!arch.length && !chiudibili.length && !depAdv.length && !attivo) return null;
          const A = derivate(somma(arch));
          return (
            <section className="card">
              <header>
                <h2>Archivio</h2>
                <span className="sp eyebrow">{arch.length} campagne fuori dai calcoli</span>
                {attivo && (
                  <button className="btn ghost mini" title="Metti da parte tutto l'advertising di questo cliente"
                    onClick={() => chiediArchiviaAdv(attivo)}>Archivia tutto l'advertising</button>
                )}
                {chiudibili.length > 0 && (
                  <button className="btn ghost mini" onClick={() => chiedi({
                    titolo: "Archiviare le campagne concluse?",
                    testo: `Sposto ${chiudibili.length} campagne in stato "conclusa" nell'archivio. I dati restano, ma escono da medie, totali, semaforo e avvisi.`,
                    etichetta: "Archivia",
                    azione: () => upd("campagne", (l) => l.map((x) => (chiudibili.some((k) => k.id === x.id) ? { ...x, archiviata: "si" } : x))),
                  })}>Archivia le {chiudibili.length} concluse</button>
                )}
                {arch.length > 0 && (
                  <button className="btn ghost mini" data-on={mostraArch ? "1" : "0"} onClick={() => setMostraArch(!mostraArch)}>
                    {mostraArch ? "Nascondi" : `Mostra le ${arch.length}`}
                  </button>
                )}
              </header>
              {depAdv.length > 0 && (
                <div className="wrap">
                  <table className="tbl">
                    <thead><tr>
                      <th>Advertising in deposito</th><th>Archiviato il</th><th className="r">Campagne</th>
                      <th className="r">Finestre</th><th className="r">Spesa storica</th>
                      <th className="r">Lead</th><th className="r">CPL</th><th />
                    </tr></thead>
                    <tbody>
                      {mapValuta(depAdv, (v2) => v2.clienteId || (v2.ids || [])[0], (v2) => {
                        const ks = v2.dati.campagne || [];
                        const pz = v2.dati.periodi || [];
                        const sp = ks.reduce((s2, k) => s2 + (k.spesa || 0), 0) + pz.reduce((s2, x) => s2 + (x.spesa || 0), 0);
                        const ld = ks.reduce((s2, k) => s2 + (k.lead || 0), 0) + pz.reduce((s2, x) => s2 + (x.lead || 0), 0);
                        return (
                          <tr key={v2.id}>
                            <td>
                              <b>{nomeCli(v2.clienteId) || "Cliente rimosso"}</b>
                              <div className="sub">{v2.etichetta === "gruppo" ? "con i sottoclienti" : "solo questo account"}</div>
                            </td>
                            <td className="sub">{dataItA(v2.quando)}</td>
                            <td className="r">{int(ks.length)}</td>
                            <td className="r">{int(pz.length)}</td>
                            <td className="r">{sp ? eur0(sp) : "—"}</td>
                            <td className="r">{ld ? int(ld) : "—"}</td>
                            <td className="r">{ld ? eur(sp / ld) : "—"}</td>
                            <td className="r"><button className="btn ghost mini" onClick={() => chiediRiattivaAdv(v2)}>Riattiva</button></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}

              {arch.length === 0 && chiudibili.length === 0 && depAdv.length === 0 && attivo && (
                <div className="nota" style={{ paddingTop: 14 }}>
                  Qui finiscono le campagne che metti da parte. Con "Archivia tutto l'advertising" sposti in deposito campagne,
                  ad set, finestre e rilevazioni di questo cliente in un colpo solo: funnel, ciclo di vendita, chiusure e task restano intatti.
                </div>
              )}

              {arch.length > 0 && !mostraArch && (
                <div className="nota" style={{ paddingTop: 14 }}>
                  In archivio ci sono {arch.length} campagne, {eur(A.sp)} di spesa storica, {int(A.cv)} conversioni e un ROAS di {A.rc ? `${fmt(A.roas)}×` : "—"}.
                  Questi numeri non entrano da nessuna parte. Premi "Mostra le {arch.length}" per vederle: ogni riga ha il pulsante Ripristina, che la rimette in elenco e nei calcoli.
                </div>
              )}
              {mostraArch && arch.length > 0 && (
                <div className="wrap">
                  <table className="tbl">
                    <thead><tr>
                      <th>Campagna</th><th>Periodo</th><th className="r">Speso</th><th className="r">CTR link</th>
                      <th className="r">Lead</th><th className="r">CPL</th>
                      <th className="r">Conv.</th><th className="r">CPA</th><th className="r">ROAS</th><th />
                    </tr></thead>
                    <tbody>
                      {mapValuta(arch, (k) => k.clienteId, (k) => (
                        <tr key={k.id} className="clic" onClick={() => apri("campagne", data.campagne.find((x) => x.id === k.id))}>
                          <td><b>{k.nome}</b><div className="sub">{et(PIATTAFORME, k.piattaforma)}{attivo ? "" : ` · ${nomeCli(k.clienteId)}`}</div></td>
                          <td className="sub">{dataIt(k.inizio)}–{dataIt(k.fine)}</td>
                          <td className="r">{eur(k.d.sp)}</td>
                          <td className="r">{pctv(k.d.ctrLink || k.d.ctr)}</td>
                          <td className="r">{k.d.ld ? int(k.d.ld) : "—"}</td>
                          <td className="r">{k.d.ld ? eur(k.d.cpl) : "—"}</td>
                          <td className="r">{int(k.d.cv)}</td>
                          <td className="r">{k.d.cv ? eur(k.d.cpa) : "—"}</td>
                          <td className="r">{k.d.rc ? `${fmt(k.d.roas)}×` : "—"}</td>
                          <td className="r">
                            <button className="btn ghost mini" onClick={(e) => { e.stopPropagation(); upd("campagne", (l) => l.map((x) => (x.id === k.id ? { ...x, archiviata: "no" } : x))); }}>Ripristina</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot><tr className="totale">
                      <td><b>Totale archivio</b></td><td />
                      <td className="r"><b>{eur(A.sp)}</b></td>
                      <td className="r"><b>{pctv(A.ctrLink)}</b></td>
                      <td className="r"><b>{A.ld ? int(A.ld) : "—"}</b></td>
                      <td className="r"><b>{A.ld ? eur(A.cpl) : "—"}</b></td>
                      <td className="r"><b>{int(A.cv)}</b></td>
                      <td className="r"><b>{A.cv ? eur(A.cpa) : "—"}</b></td>
                      <td className="r"><b>{A.rc ? `${fmt(A.roas)}×` : "—"}</b></td>
                      <td />
                    </tr></tfoot>
                  </table>
                </div>
              )}
            </section>
          );
        })()}

        {mescolate && !fAdv.linea && (
          <article className="al" data-g="nota">
            <div className="meta">
              <span className="tag">{perLinea.filter((x) => x.id !== "__senza").length} linee</span>
              <span className="misura">i numeri qui sopra le sommano</span>
            </div>
            <h3>Questo cliente lavora su più linee di business</h3>
            <dl>
              <dt>Perché conta</dt>
              <dd>CPL, CPA e ROAS medi su linee con offerte e margini diversi non descrivono nessuna delle due: una buona nasconde una cattiva, e la soglia di valutazione perde senso perché il valore medio dell'ordine cambia.</dd>
              <dt>Cosa fare</dt>
              <dd>Usa il menu delle linee in cima per isolarne una, oppure leggi la tabella qui sotto, dove restano separate.</dd>
            </dl>
          </article>
        )}

        {perLinea.length > 1 && (
          <section className="card">
            <header>
              <h2>Totali per linea di business</h2>
              <span className="sp eyebrow">tenuti separati, non sommati</span>
            </header>
            <div className="wrap">
              <table className="tbl">
                <thead><tr>
                  <th>Linea</th><th className="r">Campagne</th><th className="r">Speso</th><th className="r">CTR link</th>
                  <th className="r">Lead</th><th className="r">CPL</th><th className="r">Conv.</th><th className="r">CPA</th>
                  <th className="r">ROAS</th><th className="r">Contratti</th>
                </tr></thead>
                <tbody>
                  {perLinea.map((g) => (
                    <tr key={g.id} className="clic" onClick={() => setFAdv({ ...fAdv, linea: g.id })}>
                      <td>
                        <b>{g.nome}</b>
                        {g.id === "__senza" && <div className="sub">campagne non assegnate a una linea</div>}
                      </td>
                      <td className="r">{g.n}</td>
                      <td className="r">{eur0(g.sp)}</td>
                      <td className="r">{pctv(g.ctrLink)}</td>
                      <td className="r">{g.ld ? int(g.ld) : "—"}</td>
                      <td className="r">{g.ld ? eur(g.cpl) : "—"}</td>
                      <td className="r">{int(g.cv)}</td>
                      <td className="r">{g.cv ? eur(g.cpa) : "—"}</td>
                      <td className="r">{g.rc ? `${fmt(g.roas)}×` : "—"}</td>
                      <td className="r">{g.contratto ? eur0(g.contratto) : "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="nota">
              Qui non c'è la riga del totale, ed è voluto: sommare linee con offerte, margini e cicli di vendita diversi
              produce medie che non descrivono nessuna delle due. Clicca una riga per isolare quella linea in tutta la sezione.
            </div>
          </section>
        )}

        <section className="card">
          <header><h2>Totali per funnel</h2><span className="sp eyebrow">medie ricalcolate dai totali, non mediate tra campagne</span></header>
          <div className="wrap">
            <table className="tbl">
              <thead><tr>
                <th>Funnel</th><th className="r">Campagne</th><th className="r">Speso</th><th className="r">CTR link</th>
                <th className="r">Lead</th><th className="r">CPL</th><th className="r">Conv.</th><th className="r">CPA</th><th className="r">ROAS</th>
                <th className="r">Contratto</th><th className="r">Incassato</th>
              </tr></thead>
              <tbody>
                {perFunnel.map((g) => (
                  <tr key={g.id} className={g.id !== "__senza" ? "clic" : ""}
                    onClick={() => { if (g.id !== "__senza") { setFunnelAperto(g.id); setVista("funnel"); } }}>
                    <td><b>{g.nome}</b>{g.id === "__senza" && <div className="sub">campagne non collegate a un percorso</div>}</td>
                    <td className="r">{g.n}</td>
                    <td className="r">{eur(g.sp)}</td>
                    <td className="r">{pctv(g.ctrLink)}</td>
                    <td className="r">{g.ld ? int(g.ld) : "—"}</td>
                    <td className="r">{g.ld ? eur(g.cpl) : "—"}</td>
                    <td className="r">{int(g.cv)}</td>
                    <td className="r">{g.cv ? eur(g.cpa) : "—"}</td>
                    <td className="r">{g.rc ? `${fmt(g.roas)}×` : "—"}</td>
                    <td className="r">{g.contratto ? eur0(g.contratto) : "—"}</td>
                    <td className={"r " + (g.contratto && g.incassato < g.contratto ? "neg" : "")}>{g.contratto ? eur0(g.incassato) : "—"}</td>
                  </tr>
                ))}
              </tbody>
              {perFunnel.length > 0 && (
                <tfoot><tr className="totale">
                  <td><b>Totale</b></td>
                  <td className="r"><b>{T.n}</b></td>
                  <td className="r"><b>{eur(T.sp)}</b></td>
                  <td className="r"><b>{pctv(T.ctrLink)}</b></td>
                  <td className="r"><b>{T.ld ? int(T.ld) : "—"}</b></td>
                  <td className="r"><b>{T.ld ? eur(T.cpl) : "—"}</b></td>
                  <td className="r"><b>{int(T.cv)}</b></td>
                  <td className="r"><b>{T.cv ? eur(T.cpa) : "—"}</b></td>
                  <td className="r"><b>{T.rc ? `${fmt(T.roas)}×` : "—"}</b></td>
                  <td className="r"><b>{soldiTot(base).contratto ? eur0(soldiTot(base).contratto) : "—"}</b></td>
                  <td className="r"><b>{soldiTot(base).contratto ? eur0(soldiTot(base).incassato) : "—"}</b></td>
                </tr></tfoot>
              )}
            </table>
          </div>
          {perFunnel.some((g) => g.id === "__senza") && (
            <div className="nota">
              Le campagne senza funnel non entrano in nessun percorso: aprile e scegli la destinazione, altrimenti quando un funnel va male non sai quale traffico stava alimentando.
            </div>
          )}
        </section>

        <section className="card">
          <header>
            <h2>Copertura del messaggio</h2>
            <span className="sp eyebrow">spesa per fase del percorso d'acquisto</span>
          </header>
          {(() => {
            const attivi = creativi.filter((c) => ["attiva", "test"].includes(c.stato));
            const perStadio = STADI.map((st) => ({
              ...st, sp: attivi.filter((c) => c.stadio === st.v).reduce((s, c) => s + c.d.sp, 0),
              n: attivi.filter((c) => c.stadio === st.v).length,
            }));
            const senza = attivi.filter((c) => !c.stadio);
            const spesaNota = perStadio.reduce((s, x) => s + x.sp, 0);
            const corto = spesaNota > 0 && perStadio[0].sp + perStadio[3].sp > 0 && perStadio[1].sp + perStadio[2].sp === 0;
            if (!attivi.length) return <div className="empty"><b>Nessun ad set attivo</b>Carica ad set e annunci per vedere quali fasi del percorso stai coprendo.</div>;
            return (
              <>
                <Barre dati={perStadio.map((x) => ({ k: `${x.b}${x.n ? ` · ${x.n}` : ""}`, v: x.sp }))} fmtv={fissaValuta(eur)} />
                {senza.length > 0 && (
                  <div className="nota">{senza.length} tra ad set e annunci non hanno una fase assegnata: aprili e indica a quale momento del percorso parlano, altrimenti questo quadro resta parziale.</div>
                )}
                {corto && (
                  <article className="al" data-g="critica" style={{ margin: "0 16px 16px" }}>
                    <div className="meta"><span className="tag" data-k="late">messaggio corto</span></div>
                    <h3>Stai portando le persone dal problema all'acquisto senza passaggi intermedi</h3>
                    <dl>
                      <dt>Cosa succede</dt><dd>Tutta la spesa è concentrata sulle fasi di riconoscimento del problema e di acquisto: non c'è nulla che alimenti la raccolta di informazioni e la valutazione delle soluzioni.</dd>
                      <dt>Perché conta</dt><dd>È la causa più comune del logoramento rapido delle campagne: chi non ha ancora confrontato le alternative viene spinto a comprare, non compra, e il pubblico si consuma senza aver prodotto vendite.</dd>
                      <dt>Cosa fare</dt><dd>Aggiungi annunci per le fasi 2 e 3 — contenuti che spiegano il problema e mostrano il confronto tra le soluzioni — prima di aumentare il budget su quelli di acquisto.</dd>
                    </dl>
                  </article>
                )}
              </>
            );
          })()}
        </section>

        <section className="card">
          <header>
            <h2>Ad set e annunci</h2>
            <span className="sp eyebrow">contratti = chiusure con lo stesso nome</span>
            <button className="btn ghost mini" onClick={() => apri("creativi")}>+ Ad set o annuncio</button>
          </header>
          <div className="wrap">
            <table className="tbl">
              <thead><tr>
                <th>Nome</th><th>Livello</th><th>Fase</th><th>Campagna</th><th className="r">Speso</th><th className="r">CTR link</th>
                <th className="r">CPC link</th><th className="r">Freq.</th><th className="r">Lead</th><th className="r">CPL</th>
                <th className="r">Conv.</th><th className="r">CPA</th><th className="r">Contratti</th>
              </tr></thead>
              <tbody>
                {(() => {
                  /* Ordine gerarchico: campagna, poi ogni ad set seguito dai suoi annunci.
                     Gli annunci non collegati a un ad set restano in fondo alla loro campagna. */
                  const perCampagna = new Map();
                  creativi.forEach((c) => {
                    const k = c.campagnaId || "__senza";
                    perCampagna.set(k, [...(perCampagna.get(k) || []), c]);
                  });
                  const ordinati = [];
                  [...perCampagna.entries()]
                    .sort((a, b) => nomeCampagna(a[0]).localeCompare(nomeCampagna(b[0]), "it"))
                    .forEach(([, lista]) => {
                      const adset = lista.filter((c) => c.tipo === "adset").sort((a, b) => a.nome.localeCompare(b.nome, "it"));
                      const annunci = lista.filter((c) => c.tipo !== "adset");
                      adset.forEach((a) => {
                        ordinati.push({ ...a, liv: 0 });
                        annunci.filter((x) => x.adsetId === a.id)
                          .sort((x, y) => x.nome.localeCompare(y.nome, "it"))
                          .forEach((x) => ordinati.push({ ...x, liv: 1 }));
                      });
                      /* Gli annunci senza ad set restano a sinistra: non sono figli di nessuno */
                      annunci.filter((x) => !x.adsetId || !adset.some((a) => a.id === x.adsetId))
                        .sort((x, y) => x.nome.localeCompare(y.nome, "it"))
                        .forEach((x) => ordinati.push({ ...x, liv: 0, orfano: true }));
                    });
                  return ordinati;
                })().map((c) => (
                  <tr key={c.id} className="clic" onClick={() => apri("creativi", data.creativi.find((x) => x.id === c.id))}>
                    <td style={{ paddingLeft: 14 + c.liv * 22 }}>
                      <div className="riga" style={{ gap: 7 }}>
                        {c.liv > 0 && <span className="sub" style={{ marginLeft: -14 }}>↳</span>}
                        <div>
                          <b>{c.nome}</b>
                          {c.note && <div className="sub">{c.note}</div>}
                          {c.orfano && <div className="sub neg">da collegare a un ad set</div>}
                        </div>
                      </div>
                    </td>
                    <td><span className="tag" data-k={c.tipo === "annuncio" ? "qualificato" : ""}>{et(TIPI_CREATIVO, c.tipo)}</span></td>
                    <td className="sub">{c.stadio ? (STADI.find((x) => x.v === c.stadio) || {}).b : <span className="neg">da indicare</span>}</td>
                    <td className="sub">{nomeCampagna(c.campagnaId)}{!attivo && <div className="sub">{nomeCli(c.clienteId)}</div>}</td>
                    <td className="r">{eur(c.d.sp)}</td>
                    <td className="r">{pctv(c.d.ctrLink || c.d.ctr)}</td>
                    <td className="r">{eur(c.d.cpcLink || c.d.cpc)}</td>
                    <td className={"r " + (c.d.freq > 3.5 ? "neg" : "")}>{fmt(c.d.freq)}</td>
                    <td className="r">{c.d.ld ? int(c.d.ld) : "—"}</td>
                    <td className={"r " + (c.tetto && c.d.ld && c.d.cpl > c.tetto ? "neg" : "")}>{c.d.ld ? eur(c.d.cpl) : "—"}</td>
                    <td className="r">{int(c.d.cv)}</td>
                    <td className="r">{c.d.cv ? eur(c.d.cpa) : "—"}</td>
                    <td className={"r " + (c.contratti ? "pos" : "")}>{c.contratti ? eur0(c.contratti) : "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {!creativi.length && (
            <div className="empty"><b>Nessun ad set caricato</b>È il livello a cui ottimizzi davvero: scrivi i nomi esattamente come nella piattaforma e come nelle schede chiusura, così i contratti si agganciano da soli.</div>
          )}
        </section>
      </>
    ));
  };

  const SezioneAutomazioni = () => {
    const descriviQuando = (a) => {
      if (a.trigger === "avviso") return `Avviso${a.filtroAvviso ? ` su ${a.filtroAvviso}` : ""}${a.soloCritici === "si" ? ", solo critici" : ""}`;
      if (a.trigger === "metrica") return `${et(METRICHE_TRIGGER, a.metrica)} ${et(OPERATORI, a.operatore)} ${fmt(a.valore)}`;
      if (a.trigger === "budget") return `Erogato oltre il ${fmt(a.valore, 0)}%`;
      if (a.trigger === "task") return "Task oltre la scadenza";
      if (a.trigger === "chiusura") return "Nuova chiusura registrata";
      return `Dal giorno ${int(a.giorno)} senza dati commerciali`;
    };
    const descriviCosa = (a) => [
      a.azTask === "si" ? `task "${a.taskTitolo || "senza titolo"}" (${et(PRIORITA, a.taskPriorita)})` : null,
      a.azWebhook === "si" ? (a.url ? "webhook" : "webhook senza URL") : null,
    ].filter(Boolean).join(" · ") || "solo registro";

    /* Il pacchetto si installa sul cliente che scegli, o su tutti.
       Le regole già presenti con lo stesso nome non si duplicano. */
    const installaADA = (cid) => {
      const esistenti = new Set(data.automazioni.map((a) => `${a.nome}|${a.clienteId || ""}`));
      const nuove = REGOLE_ADA
        .filter((r) => !esistenti.has(`${r.nome}|${cid || ""}`))
        .map((r) => ({
          ...VUOTO.automazioni, id: uid(), nome: r.nome, trigger: r.trigger,
          filtroAvviso: r.filtroAvviso, soloCritici: r.soloCritici, clienteId: cid || "",
          attiva: "si", azTask: "si", taskTitolo: r.taskTitolo, taskPriorita: r.taskPriorita,
          taskGiorni: r.taskGiorni, taskArea: (data.aree[0] || {}).id || "", azWebhook: "no", url: "",
        }));
      if (!nuove.length) {
        chiedi({ titolo: "Ci sono già", testo: `Le ${REGOLE_ADA.length} regole del framework esistono già per ${cid ? nomeCli(cid) : "tutti i clienti"}. Se vuoi rifarle, elimina prima quelle presenti.`, etichetta: "Ho capito", azione: () => {} });
        return;
      }
      chiedi({
        titolo: `Installare le regole del framework?`,
        testo: `Creo ${nuove.length} regole su ${cid ? nomeCli(cid) : "tutti i clienti"}, una per ciascuna soglia: gancio, asta, conversioni, margine, CPL, CPA, frequenza, erogazione, più una rete di sicurezza sui critici.\n\nOgni regola crea un task quando l'avviso corrispondente si accende. Le soglie restano quelle della singola campagna, prese dai suoi benchmark: qui non c'è nessun numero fisso.\n\nRicorda che le automazioni girano solo mentre l'app è aperta.`,
        etichetta: `Crea le ${nuove.length} regole`,
        azione: () => upd("automazioni", (l) => [...nuove, ...l]),
      });
    };

    const prova = async (a) => {
      const ok = a.azWebhook === "si" && a.url
        ? await inviaWebhook(a.url, { automazione: a.nome, evento: "Invio di prova", dettaglio: "Test manuale dall'app", quando: new Date().toISOString() })
        : false;
      upd("registro", (l) => [{
        id: uid(), quando: new Date().toISOString(), automazione: a.nome, testo: "Prova manuale",
        esito: a.azWebhook !== "si" ? "nessun webhook configurato" : !a.url ? "URL mancante" : ok ? "richiesta partita" : "invio non riuscito",
      }, ...l].slice(0, 80));
    };

    return (
      <>
        <section className="card">
          <header>
            <h2>Framework delle decisioni</h2>
            <span className="sp eyebrow">
              {REGOLE_ADA.filter((r) => data.automazioni.some((a) => a.nome === r.nome)).length} di {REGOLE_ADA.length} installate
            </span>
            <select value={adaCliente} onChange={(e) => setAdaCliente(e.target.value)} aria-label="Su quale account">
              <option value="">Su tutti i clienti</option>
              {ordinaGerarchia().map(({ z, liv }) => (
                <option key={z.id} value={z.id}>{liv ? `${"\u00A0".repeat(liv * 4)}↳ ` : ""}{z.nome}</option>
              ))}
            </select>
            <button className="btn" onClick={() => installaADA(adaCliente)}>Installa le regole</button>
          </header>
          <div className="wrap">
            <table className="tbl">
              <thead><tr><th>Soglia</th><th>Quando scatta</th><th>Cosa crea</th><th className="r">Stato</th></tr></thead>
              <tbody>
                {REGOLE_ADA.map((r) => {
                  const dove = data.automazioni.filter((a) => a.nome === r.nome);
                  return (
                    <tr key={r.nome}>
                      <td><b>{r.nome.replace("ADA · ", "")}</b><div className="sub">{r.perche}</div></td>
                      <td className="sub">{r.filtroAvviso ? `avviso su ${r.filtroAvviso}` : "qualunque avviso critico"}</td>
                      <td className="sub">task {et(PRIORITA, r.taskPriorita).toLowerCase()} a {r.taskGiorni} giorni</td>
                      <td className="r">
                        {dove.length
                          ? <span className="tag" data-k="attiva">{dove.some((a) => !a.clienteId) ? "tutti" : `${dove.length} client${dove.length === 1 ? "e" : "i"}`}</span>
                          : <span className="sub">—</span>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div className="nota">
            Le soglie non sono scritte qui: ogni regola si appoggia agli avvisi, che usano i benchmark della singola
            campagna. Impostando un CTR minimo diverso su un cliente, la stessa regola scatta a un valore diverso.
            Senza benchmark impostati sulla campagna, l'avviso non si accende e la regola resta muta.
          </div>
        </section>

        <section className="card">
          <header>
            <h2>Regole</h2>
            <span className="sp eyebrow">{data.automazioni.filter((a) => a.attiva === "si").length} attive · girano mentre l'app è aperta</span>
          </header>
          <div className="wrap">
            <table className="tbl">
              <thead><tr><th>Regola</th><th>Quando</th><th>Cosa fa</th><th>Su</th><th>Stato</th><th className="r">Prova</th></tr></thead>
              <tbody>
                {data.automazioni.map((a) => (
                  <tr key={a.id} className="clic" onClick={() => apri("automazioni", a)}>
                    <td><b>{a.nome}</b><div className="sub">{et(TRIGGER, a.trigger)}</div></td>
                    <td className="sub">{descriviQuando(a)}</td>
                    <td className="sub">{descriviCosa(a)}</td>
                    <td className="sub">{a.clienteId ? nomeCli(a.clienteId) : "Tutti"}</td>
                    <td><span className="tag" data-k={a.attiva === "si" ? "attivo" : ""}>{a.attiva === "si" ? "attiva" : "in pausa"}</span></td>
                    <td className="r">
                      <button className="btn ghost mini" onClick={(e) => { e.stopPropagation(); prova(a); }}>Invia prova</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {!data.automazioni.length && (
            <div className="empty">
              <b>Nessuna regola</b>
              Una regola collega un innesco (una soglia, un ritardo, una chiusura) a un'azione.
              <div style={{ marginTop: 14 }}>
                <button className="btn" onClick={() => upd("automazioni", (l) => [...REGOLE_BASE(), ...l])}>
                  Inserisci le quattro regole di esempio
                </button>
              </div>
            </div>
          )}
        </section>

        <section className="card">
          <header><h2>Registro</h2><span className="sp" />
            {data.registro.length > 0 && <button className="btn ghost mini" onClick={() => upd("registro", () => [])}>Svuota</button>}
          </header>
          <div className="acts">
            {data.registro.slice(0, 25).map((r) => (
              <div className="tkrow" key={r.id}>
                <div className="t">
                  <b>{r.testo}</b>
                  <div className="meta" style={{ marginTop: 5 }}>
                    <span className="pill">{r.automazione}</span>
                    <span className="pill">{r.esito}</span>
                  </div>
                </div>
                <span className="pill">{new Date(r.quando).toLocaleString(LOCALE, { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" })}</span>
              </div>
            ))}
            {!data.registro.length && <div className="empty"><b>Nessuno scatto</b>Qui finisce ogni regola che parte, con l'esito.</div>}
          </div>
        </section>

        <section className="card">
          <header><h2>Come funziona, e cosa non può fare</h2></header>
          <div style={{ padding: "14px 16px", fontSize: 13.3, color: "var(--ink2)" }}>
            Le regole vengono controllate a ogni cambiamento dei dati, quindi mentre l'app è aperta. Ogni evento parte una volta
            sola: lo stesso avviso sulla stessa campagna non ricrea un task ogni volta che apri la sezione.
            <br /><br />
            I webhook escono in uscita: puoi chiamare uno scenario Make, uno Zap, un flusso n8n o un endpoint tuo, e da lì fare
            il resto — messaggio in Slack, riga in un foglio, email al cliente. Il browser però non può leggere la risposta,
            quindi la conferma di consegna la vedi solo sul servizio ricevente. In entrata, invece, non si può ricevere nulla:
            per far arrivare un lead da un modulo Meta dentro questa app servirebbe un server sempre acceso con un indirizzo pubblico.
          </div>
        </section>
      </>
    );
  };

  const SezioneChiusure = () => {
    const righe = [...chiusureVista].sort((a, b) => (a.data < b.data ? 1 : -1));
    /* I conti seguono la stessa regola del Riepilogo mensile: in un mese con schede chiusura
       valgono le schede, altrimenti i numeri del Ciclo di vendita (scritti a mano, dal registro
       dei venditori o da un file KPI). Così la sezione si aggiorna da sola quando cambia il ciclo. */
    const ids = attivo ? ambito : attivi.map((z) => z.id);
    const mesiCh = righeMensili(ids).filter((r) => r.chiusure || r.contratto || r.incassato).reverse();
    const nChiusure = mesiCh.reduce((s, r) => s + r.chiusure, 0);
    const contratto = mesiCh.reduce((s, r) => s + r.contratto, 0);
    const incassato = mesiCh.reduce((s, r) => s + r.incassato, 0);
    const daCiclo = mesiCh.some((r) => !data.chiusure.some((c) => ids.includes(c.clienteId) && mese(c.data) === r.mm));
    const nomeCampagna = (id) => (data.campagne.find((k) => k.id === id) || {}).nome || "—";
    return conValuta(valutaComune(righe.map((c) => c.clienteId)), () => (
      <>
        <div className="kpis">
          <div className="kpi"><IcoKpi t="Clienti chiusi" /><div className="eyebrow">Clienti chiusi</div><div className="v">{int(nChiusure)}</div><div className="s">contratto medio {nChiusure ? eur(contratto / nChiusure) : "—"}{daCiclo ? " · anche dal ciclo di vendita" : ""}</div></div>
          <div className="kpi"><IcoKpi t="Valore contratto" /><div className="eyebrow">Valore contratto</div><div className="v">{eur0(contratto)}</div><div className="s">somma degli accordi firmati</div></div>
          <div className="kpi"><IcoKpi t="Incassato" /><div className="eyebrow">Incassato</div><div className="v">{eur0(incassato)}</div><div className="s">{contratto ? pctv((incassato / contratto) * 100) : "—"} del contrattualizzato</div></div>
          <div className="kpi"><IcoKpi t="Da incassare" /><div className="eyebrow">Da incassare</div><div className="v" style={{ color: contratto - incassato ? "var(--acc)" : "var(--ink)" }}>{eur0(contratto - incassato)}</div><div className="s">saldi ancora aperti</div></div>
        </div>

        {!attivo && (() => {
          /* Fuori dai clienti solo i totali per account: le singole schede sono dentro il cliente */
          const perCli = righeGruppi("chiusure", (ids) => {
            const l = righeMensili(ids).filter((r) => r.chiusure || r.contratto || r.incassato);
            if (!l.length) return null;
            const sc = righe.filter((c) => ids.includes(c.clienteId));
            return { n: l.reduce((x, r) => x + r.chiusure, 0), ct: l.reduce((x, r) => x + r.contratto, 0), ic: l.reduce((x, r) => x + r.incassato, 0),
              ultima: l[l.length - 1].mm, schede: sc.length };
          });
          return (
            <section className="card">
              <header><h2>Chiusure per cliente</h2><span className="sp eyebrow">totali · ▸ apre i sottoclienti · le singole schede sono dentro ogni cliente</span></header>
              <div className="wrap"><table className="tbl">
                <thead><tr><th>Account</th><th className="r">Chiusure</th><th className="r">Contratto</th><th className="r">Incassato</th><th className="r">Residuo</th><th className="r">Ultimo mese</th></tr></thead>
                <tbody>
                  {mapValuta(perCli, (r) => r.z.id, (r) => (
                    <tr key={r.key} className="clic" onClick={() => { cambiaAccount(r.z.id); setVista("chiusure"); }} title="Apri le chiusure del cliente">
                      <td style={{ paddingLeft: 14 + r.liv * 18 }}>{nomeGruppo("chiusure", r)}</td>
                      <td className="r">{int(r.n)}<div className="sub">{r.schede ? `${int(r.schede)} schede` : "dal ciclo di vendita"}</div></td>
                      <td className="r">{eur0(r.ct)}</td>
                      <td className="r">{eur0(r.ic)}</td>
                      <td className={"r " + (r.ct - r.ic > 0 ? "neg" : "")}>{r.ct - r.ic > 0 ? eur0(r.ct - r.ic) : "saldato"}</td>
                      <td className="r sub">{meseIt(r.ultima)}</td>
                    </tr>
                  ))}
                </tbody>
              </table></div>
              {!perCli.length && <div className="empty"><b>Nessuna chiusura registrata</b>Arrivano dalle schede chiusura e dal Ciclo di vendita di ogni cliente.</div>}
            </section>
          );
        })()}

        {attivo && mesiCh.length > 0 && (
          <section className="card">
            <header><h2>Chiusure mese per mese</h2><span className="sp eyebrow">dalle schede quando ci sono, altrimenti dal Ciclo di vendita</span></header>
            <div className="wrap"><table className="tbl">
              <thead><tr><th>Mese</th><th className="r">Chiusure</th><th className="r">Contratto</th><th className="r">Incassato</th><th className="r">Residuo</th><th className="r">Contratto medio</th><th className="r">Costo per chiusura</th></tr></thead>
              <tbody>
                {mesiCh.map((r) => {
                  const schede = righe.filter((c) => mese(c.data) === r.mm).length;
                  return (
                    <tr key={r.mm}>
                      <td><b>{meseIt(r.mm)}</b><div className="sub">{schede ? `${schede} schede` : "dal ciclo di vendita"}</div></td>
                      <td className="r">{int(r.chiusure)}</td>
                      <td className="r">{eur0(r.contratto)}</td>
                      <td className="r">{eur0(r.incassato)}</td>
                      <td className={"r " + (r.contratto - r.incassato > 0 ? "neg" : "")}>{r.contratto - r.incassato > 0 ? eur0(r.contratto - r.incassato) : "saldato"}</td>
                      <td className="r">{r.chiusure ? eur(r.contratto / r.chiusure) : "—"}</td>
                      <td className="r">{r.chiusure && r.spesa ? eur(r.spesa / r.chiusure) : "—"}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table></div>
          </section>
        )}

        {attivo && <section className="card">
          <header><h2>Clienti chiusi</h2><span className="sp eyebrow">ogni riga alimenta il mese corrispondente</span></header>
          <div className="wrap">
            <table className="tbl">
              <thead><tr>
                <th>Cliente chiuso</th>{!attivo && <th>Account</th>}
                <th>Provenienza</th><th>Ad set e annuncio</th>
                <th className="r">Contratto</th><th className="r">Incassato</th><th className="r">Residuo</th><th className="r">Data</th>
              </tr></thead>
              <tbody>
                {mapValuta(righe, (c) => c.clienteId, (c) => {
                  const res = (c.valoreContratto || 0) - (c.incassato || 0);
                  return (
                    <tr key={c.id} className="clic" onClick={() => apri("chiusure", c)}>
                      <td><b>{c.nome}</b><div className="sub">{[c.azienda, c.email].filter(Boolean).join(" · ")}</div></td>
                      {!attivo && <td className="sub">{nomeCli(c.clienteId)}</td>}
                      <td><span className="tag" data-k="qualificato">{et(FUNNEL_PROV, c.funnel)}</span><div className="sub">{nomeCampagna(c.campagnaId)}</div></td>
                      <td className="sub">{c.adset || "—"}<br />{c.annuncio || "—"}</td>
                      <td className="r">{eur(c.valoreContratto)}</td>
                      <td className="r">{eur(c.incassato)}</td>
                      <td className={"r " + (res > 0 ? "neg" : "")}>{res > 0 ? eur(res) : "saldato"}</td>
                      <td className="r sub">{dataIt(c.data)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          {!righe.length && (
            <div className="empty"><b>Nessuna scheda chiusura</b>{mesiCh.length ? "I numeri qui sopra arrivano dal Ciclo di vendita. Le schede servono quando vuoi sapere da quale campagna, ad set e annuncio è arrivato ogni cliente chiuso." : "Ogni scheda tiene insieme il cliente chiuso, gli importi e l'annuncio esatto da cui è arrivato."}</div>
          )}
        </section>}
      </>
    ));
  };

  const SezioneVendite = () => {
    const righe = [...venditeVista].sort((a, b) => (a.periodo < b.periodo ? 1 : -1));
    const t = sommaVendite(venditeVista);
    return conValuta(valutaComune(venditeVista.map((v) => v.clienteId)), () => (
      <>
        <div className="kpis">
          <div className="kpi"><IcoKpi t="Chiusure" /><div className="eyebrow">Chiusure</div><div className="v">{int(t.chiusure)}</div><div className="s">su {int(t.demoSvolte)} demo svolte</div></div>
          <div className="kpi"><IcoKpi t="Tasso di chiusura" /><div className="eyebrow">Tasso di chiusura</div><div className="v">{t.demoSvolte ? pctv((t.chiusure / t.demoSvolte) * 100) : "—"}</div><div className="s">lead → chiusura {t.lead ? pctv((t.chiusure / t.lead) * 100) : "—"}</div></div>
          <div className="kpi"><IcoKpi t="Valore contratto" /><div className="eyebrow">Valore contratto</div><div className="v">{eur0(t.valoreContratto)}</div><div className="s">contratto medio {t.chiusure ? eur(t.valoreContratto / t.chiusure) : "—"}</div></div>
          <div className="kpi"><IcoKpi t="Incassato" /><div className="eyebrow">Incassato</div><div className="v">{eur0(t.incassato)}</div><div className="s">{t.valoreContratto ? `${pctv((t.incassato / t.valoreContratto) * 100)} del contratto` : "nessun contratto"}</div></div>
        </div>

        {(t.slot > 0 || t.oreContatto > 0) && (
          <div className="kpis">
            <div className="kpi">
              <IcoKpi t="Uso della capacità" />
              <div className="eyebrow">Uso della capacità</div>
              <div className="v" style={{ color: t.slot && t.discoveryFissate / t.slot > 0.9 ? "var(--acc)" : "var(--ink)" }}>
                {t.slot ? pctv((t.discoveryFissate / t.slot) * 100, 0) : "—"}
              </div>
              <div className="s">{int(t.discoveryFissate)} discovery su {int(t.slot)} slot disponibili</div>
            </div>
            <div className="kpi">
              <IcoKpi t="Tempo al primo contatto" />
              <div className="eyebrow">Tempo al primo contatto</div>
              <div className="v" style={{ color: venditeVista.length && (t.oreContatto / venditeVista.length) > 12 ? "var(--acc)" : "var(--ink)" }}>
                {venditeVista.length ? `${fmt(t.oreContatto / venditeVista.length, 1)} h` : "—"}
              </div>
              <div className="s">media dei periodi caricati</div>
            </div>
            <div className="kpi"><IcoKpi t="Lead per slot" /><div className="eyebrow">Lead per slot</div><div className="v">{t.slot ? fmt(t.lead / t.slot, 1) : "—"}</div><div className="s">quanti lead per ogni posto disponibile</div></div>
            <div className="kpi"><IcoKpi t="Lead non lavorati" /><div className="eyebrow">Lead non lavorati</div><div className="v">{int(Math.max(0, t.lead - t.chiamate))}</div><div className="s">lead ricevuti meno chiamate fatte</div></div>
          </div>
        )}

        {venditeVista.length > 0 && t.oreContatto / venditeVista.length > 12 && (
          <article className="al" data-g="attenzione">
            <div className="meta"><span className="tag">collo di bottiglia</span>
              <span className="misura">primo contatto <b>{fmt(t.oreContatto / venditeVista.length, 1)} h</b> · obiettivo &lt; 12 h</span></div>
            <h3>I lead aspettano troppo prima della prima chiamata</h3>
            <dl>
              <dt>Cosa succede</dt><dd>In media passano {fmt(t.oreContatto / venditeVista.length, 1)} ore tra l'arrivo del lead e la prima chiamata. In quel tempo l'interesse cala e i concorrenti rispondono prima.</dd>
              <dt>Causa probabile</dt><dd>Capacità commerciale insufficiente rispetto al volume che le campagne generano, oppure i lead non arrivano subito a chi deve chiamare.</dd>
              <dt>Cosa fare</dt><dd>Prima di aumentare il budget, riduci questo tempo: è la leva con più effetto sul tasso di chiusura e non costa un euro di spesa pubblicitaria in più.</dd>
            </dl>
          </article>
        )}

        {(() => {
          const t = sommaVendite(venditeVista);
          if (!attivo || !t.demoFissate) return null;
          const show = t.demoSvolte / t.demoFissate;
          const attivi = S.creativi.filter((c) => ["attiva", "test"].includes(c.stato));
          const medie = attivi.filter((c) => ["2", "3"].includes(c.stadio)).reduce((s, c) => s + (c.spesa || 0), 0);
          const totali = attivi.reduce((s, c) => s + (c.spesa || 0), 0);
          if (show >= 0.6 || !totali || medie / totali > 0.15) return null;
          return (
            <article className="al" data-g="critica">
              <div className="meta">
                <span className="tag" data-k="late">causa a monte</span>
                <span className="misura">show up demo <b>{pctv(show * 100, 0)}</b> · fasi 2 e 3 {pctv((medie / totali) * 100, 0)} della spesa</span>
              </div>
              <h3>Le mancate presenze nascono dal messaggio, non dal commerciale</h3>
              <dl>
                <dt>Cosa succede</dt>
                <dd>Solo il {pctv(show * 100, 0)} di chi prenota si presenta, e quasi tutta la spesa pubblicitaria è concentrata sulle fasi di problema e acquisto: chi prenota non ha mai attraversato la raccolta di informazioni né il confronto tra le soluzioni.</dd>
                <dt>Perché conta</dt>
                <dd>Portare le persone dal problema all'acquisto con un annuncio solo produce tre effetti riconoscibili: campagne che si rompono quando alzi il budget, qualità dei lead bassa, e appuntamenti presi da chi non è ancora pronto — quindi non si presenta. È lo stesso difetto visto da tre angoli.</dd>
                <dt>Cosa fare</dt>
                <dd>Prima di aggiungere promemoria e telefonate di conferma, metti in campo gli annunci delle fasi 2 e 3: la guida all'acquisto e il chi siamo che supera l'obiezione più grande. Li trovi già scritti nella sezione Messaggio.</dd>
              </dl>
            </article>
          );
        })()}

        {venditeVista.some((r) => r.convPiattaforma > 0) && (() => {
          const conf = venditeVista.filter((r) => r.convPiattaforma > 0 && r.lead > 0);
          const piat = conf.reduce((s, r) => s + r.convPiattaforma, 0);
          const reali = conf.reduce((s, r) => s + r.lead, 0);
          const prec = piat ? reali / piat : 0;
          const buona = prec >= 0.85 && prec <= 1.05;
          return (
            <article className="al" data-g={buona ? "nota" : "attenzione"}>
              <div className="meta">
                <span className="tag" data-k={buona ? "attivo" : ""}>{buona ? "dati affidabili" : "scarto da verificare"}</span>
                <span className="misura">precisione <b>{pctv(prec * 100, 0)}</b> · attesa 85–95%</span>
              </div>
              <h3>{buona ? "I numeri delle piattaforme reggono il confronto" : "Le piattaforme e i tuoi dati non raccontano la stessa cosa"}</h3>
              <dl>
                <dt>Cosa succede</dt>
                <dd>Le piattaforme dichiarano {int(piat)} conversioni, tu ne hai registrate {int(reali)} come lead reali: uno scarto del {pctv(Math.abs(1 - prec) * 100, 0)}.</dd>
                <dt>Come leggerlo</dt>
                <dd>Nessun sistema arriva al 100%, e sistemi diversi non si riconcilieranno mai del tutto. L'obiettivo non è la perfezione: è conoscere lo scarto e vederlo stabile nel tempo. Fra l'85 e il 95% si può comprare media con fiducia.</dd>
                <dt>Cosa fare</dt>
                <dd>{buona
                  ? "Continua a confrontare ogni mese: quando lo scarto cambia all'improvviso, quasi sempre è il tracciamento che si è rotto, non il mercato che è cambiato."
                  : "Prima di riscrivere le campagne, verifica il tracciamento con una conversione di prova e controlla che i dati arrivino puliti. Decidere su numeri sbagliati costa più che non decidere."}</dd>
              </dl>
            </article>
          );
        })()}

        {PannelloKpi()}

        {attivo && (() => {
          const gs = data.giornate.filter((g) => (attivo ? ambito.includes(g.clienteId) : true));
          const nomeFile = `registro-venditori-${attivo ? attivo.nome.toLowerCase().replace(/[^a-z0-9]+/g, "-") : "clienti"}-${oggi()}`;
          const intest = COL_DIARIO.map((c) => c.l);
          const esempio = COL_DIARIO.map((c) => (c.k === "data" ? oggi() : c.ex));

          const scaricaExcel = () => {
            const righe = [intest, esempio, ...Array.from({ length: 30 }, () => intest.map(() => ""))];
            const foglio = XLSX.utils.aoa_to_sheet(righe);
            foglio["!cols"] = COL_DIARIO.map((c) => ({ wch: Math.max(12, c.l.length + 3) }));
            const istruzioni = XLSX.utils.aoa_to_sheet([
              ["Registro giornaliero dei venditori"],
              [attivo ? attivo.nome : ""],
              [],
              ["Come si compila"],
              ["Una riga per ogni giorno e per ogni venditore: se in un giorno hanno lavorato in due, si scrivono due righe."],
              ["I numeri sono quelli di quel giorno soltanto, non i totali dall'inizio del mese."],
              ["La colonna Funnel indica da dove arrivavano i lead chiamati in quella giornata."],
              ["Le righe senza numeri vengono ignorate: si puo lasciare vuoto un giorno in cui non si e lavorato."],
              [],
              ["Come si consegna"],
              ["Una volta a settimana, con il file completo dall'inizio del mese: ricaricarlo non crea doppioni."],
              ["Le righe si riconoscono da data, venditore e funnel, quindi una correzione sostituisce quella precedente."],
              [],
              ["Cosa non va scritto qui"],
              ["Nomi e recapiti dei clienti chiusi: quelli si registrano nella scheda chiusura, con contratto e incassato."],
              [],
              ["Le colonne"],
              ...COL_DIARIO.map((c) => [c.l, c.k === "data" ? "formato giorno/mese/anno oppure anno-mese-giorno" : (c.ex ? `esempio: ${c.ex}` : "")]),
            ]);
            istruzioni["!cols"] = [{ wch: 34 }, { wch: 72 }];
            const cartella = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(cartella, foglio, "Registro");
            XLSX.utils.book_append_sheet(cartella, istruzioni, "Istruzioni");
            const dati = XLSX.write(cartella, { bookType: "xlsx", type: "array" });
            scaricaFile(`${nomeFile}.xlsx`, dati, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
          };

          const scaricaCsv = () => {
            const testo = [intest.join(";"), esempio.join(";"),
              ...Array.from({ length: 10 }, () => intest.map(() => "").join(";"))].join("\n");
            scaricaFile(`${nomeFile}.csv`, "\uFEFF" + testo, "text/csv;charset=utf-8");
          };

          const leggi = (file) => {
            const cid = attivo ? attivo.id : "";
            if (!cid) { chiedi({ titolo: "Scegli prima il cliente", testo: "Entra nell'account del cliente e ricarica il file da lì: le giornate vanno assegnate a un cliente preciso.", etichetta: "Ho capito", azione: () => {} }); return; }
            const applica = (testo) => {
              const { intestazioni, righe } = leggiTabella(testo);
              const pulisci = (x) => String(x || "").toLowerCase().replace(/[^a-z0-9]/g, "");
              const mappa = {};
              COL_DIARIO.forEach((c) => {
                const i = intestazioni.findIndex((h) => pulisci(h) === pulisci(c.l));
                if (i > -1) mappa[c.k] = i;
              });
              if (mappa.data === undefined) {
                /* Non è il registro: forse è un file KPI di riepilogo, lo passo al riquadro sopra */
                leggiFileKpi(file);
                setTimeout(() => { const el = document.getElementById("kpi-import"); if (el) el.scrollIntoView({ behavior: "smooth", block: "start" }); }, 100);
                return;
              }
              const nuove = [];
              righe.forEach((r) => {
                const dataG = normalizzaData(r[mappa.data]);
                if (!dataG) return;
                const rec = { id: uid(), clienteId: cid, data: dataG,
                  venditore: (r[mappa.venditore] || "").trim() || "non indicato",
                  funnel: (r[mappa.funnel] || "").trim() || "non indicato",
                  note: mappa.note !== undefined ? (r[mappa.note] || "").trim() : "" };
                NUM_DIARIO.forEach((k) => { rec[k] = mappa[k] !== undefined ? parseNumIntl(r[mappa[k]]) : 0; });
                const vuota = NUM_DIARIO.every((k) => !rec[k]);
                if (!vuota) nuove.push(rec);
              });
              if (!nuove.length) {
                chiedi({ titolo: "Nessuna riga utile", testo: "Il file non contiene giornate con dei numeri.", etichetta: "Ho capito", azione: () => {} });
                return;
              }
              const chiave = (g) => `${g.data}|${g.venditore}|${g.funnel}`;
              const chiavi = nuove.map(chiave);
              const sostituite = gs.filter((g) => g.clienteId === cid && chiavi.includes(chiave(g))).length;
              const mesi = [...new Set(nuove.map((g) => mese(g.data)))];
              chiedi({
                titolo: "Caricare il registro?",
                testo: `${nuove.length} giornate da ${[...new Set(nuove.map((g) => g.venditore))].length} venditori, su ${mesi.map(meseIt).join(", ")}.` +
                  (sostituite ? `\n\n${sostituite} righe già presenti con stessa data, venditore e funnel verranno sostituite: ricaricare lo stesso file non crea doppioni.` : ""),
                etichetta: "Carica",
                azione: () => {
                  setData((d) => {
                    const resta = d.giornate.filter((g) => !(g.clienteId === cid && chiavi.includes(chiave(g))));
                    const vendite = [...d.vendite];
                    mesi.forEach((mm) => {
                      if (!vendite.some((v) => v.clienteId === cid && v.periodo === mm)) {
                        vendite.unshift({ id: uid(), clienteId: cid, periodo: mm, lead: 0, chiamate: 0, discoveryFissate: 0,
                          discoverySvolte: 0, demoFissate: 0, demoSvolte: 0, offerte: 0, chiusure: 0, valoreContratto: 0,
                          incassato: 0, spesa: 0, slot: 0, oreContatto: 0, convPiattaforma: 0 });
                      }
                    });
                    return { ...d, giornate: [...resta, ...nuove], vendite };
                  });
                },
              });
            };
            if (eExcel(file.name)) {
              const r = new FileReader();
              r.onload = () => {
                try {
                  const f = fogliDaExcel(r.result);
                  if (!f.length) return;
                  const buono = f.find((x) => /(^|;|,|\t)\s*data\s*($|;|,|\t)/i.test(x.csv.split("\n")[0] || "")) || f[0];
                  applica(buono.csv);
                } catch (e) { console.error(e); }
              };
              r.readAsArrayBuffer(file);
            } else {
              const r = new FileReader();
              r.onload = () => applica(r.result);
              r.readAsText(file);
            }
          };

          const perChi = (campo) => {
            const mm = new Map();
            gs.forEach((g) => {
              const key = g[campo] || "non indicato";
              const x = mm.get(key) || { lead: 0, chiamate: 0, demoFissate: 0, demoSvolte: 0, chiusure: 0, valoreContratto: 0, giorni: 0 };
              NUM_DIARIO.forEach((k) => { if (x[k] !== undefined) x[k] += g[k] || 0; });
              x.giorni += 1;
              mm.set(key, x);
            });
            return [...mm.entries()].sort((a, b) => b[1].chiusure - a[1].chiusure);
          };

          return (
            <>
              <section className="card">
                <header>
                  <h2>Registro dei venditori</h2>
                  <span className="sp eyebrow">{gs.length ? `${gs.length} giornate caricate` : "una riga al giorno per venditore"}</span>
                  <button className="btn ghost mini" onClick={scaricaCsv}>CSV</button>
                  <button className="btn" onClick={scaricaExcel}>Scarica il modulo Excel</button>
                </header>
                <div style={{ padding: "16px 18px", display: "grid", gap: 10, maxWidth: 560 }}>
                  <div className="f">
                    <label htmlFor="diario-file">Carica il modulo compilato (CSV o Excel)</label>
                    <div className="drop" data-sopra={trascina ? "1" : "0"}
                      onDragOver={(e) => { e.preventDefault(); setTrascina(true); }}
                      onDragLeave={() => setTrascina(false)}
                      onDrop={(e) => {
                        e.preventDefault(); setTrascina(false);
                        const fl = e.dataTransfer.files && e.dataTransfer.files[0];
                        if (fl) leggi(fl);
                      }}>
                      <input id="diario-file" type="file" accept=".csv,.tsv,.txt,.xlsx,.xls,.ods"
                        onChange={(e) => { const fl = e.target.files && e.target.files[0]; if (fl) leggi(fl); e.target.value = ""; }} />
                      <span className="sub">…oppure trascina qui il modulo compilato</span>
                    </div>
                    <span className="hint">
                      Il modulo ha una riga per giorno, con venditore e funnel di provenienza. Le giornate caricate sostituiscono
                      i totali scritti a mano nel mese corrispondente, e ricaricare lo stesso file non crea doppioni.
                    </span>
                  </div>
                </div>
              </section>

              {gs.length > 0 && (
                <div className="two">
                  <section className="card">
                    <header><h2>Per venditore</h2><span className="sp eyebrow">dalle giornate caricate</span></header>
                    <div className="wrap">
                      <table className="tbl">
                        <thead><tr><th>Venditore</th><th className="r">Giorni</th><th className="r">Chiamate</th><th className="r">Demo svolte</th><th className="r">Show up</th><th className="r">Chiusure</th><th className="r">Chiusura</th></tr></thead>
                        <tbody>
                          {perChi("venditore").map(([nome, x]) => (
                            <tr key={nome}>
                              <td><b>{nome}</b></td>
                              <td className="r">{int(x.giorni)}</td>
                              <td className="r">{int(x.chiamate)}</td>
                              <td className="r">{int(x.demoSvolte)}</td>
                              <td className="r">{x.demoFissate ? pctv((x.demoSvolte / x.demoFissate) * 100, 0) : "—"}</td>
                              <td className="r">{int(x.chiusure)}</td>
                              <td className="r">{x.demoSvolte ? pctv((x.chiusure / x.demoSvolte) * 100, 0) : "—"}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </section>

                  <section className="card">
                    <header><h2>Per funnel di provenienza</h2><span className="sp eyebrow">dei lead chiamati</span></header>
                    <div className="wrap">
                      <table className="tbl">
                        <thead><tr><th>Funnel</th><th className="r">Lead</th><th className="r">Demo svolte</th><th className="r">Chiusure</th><th className="r">Contratto</th></tr></thead>
                        <tbody>
                          {perChi("funnel").map(([nome, x]) => (
                            <tr key={nome}>
                              <td><b>{nome}</b></td>
                              <td className="r">{int(x.lead)}</td>
                              <td className="r">{int(x.demoSvolte)}</td>
                              <td className="r">{int(x.chiusure)}</td>
                              <td className="r">{x.valoreContratto ? eur0(x.valoreContratto) : "—"}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <div className="nota">
                      Qui il funnel è quello dichiarato dal venditore su chi ha chiamato: confrontalo con la provenienza registrata
                      nelle schede chiusura, perché quando i due non coincidono di solito è il tracciamento a mentire, non il venditore.
                    </div>
                  </section>
                </div>
              )}
            </>
          );
        })()}

        {!attivo && (() => {
          /* Fuori dai clienti solo i totali per cliente: i mesi nel dettaglio sono dentro il cliente */
          const perCli = righeGruppi("vendite", (ids) => {
            const l = righe.filter((r) => ids.includes(r.clienteId));
            const periodi = [...new Set(l.map((r) => r.periodo))].sort();
            return l.length ? { mesi: periodi.length, ultimo: periodi[periodi.length - 1], ...sommaVendite(l) } : null;
          });
          return (
            <section className="card">
              <header><h2>Ciclo di vendita per cliente</h2><span className="sp eyebrow">totali dei mesi registrati · ▸ apre i sottoclienti · il dettaglio è dentro ogni cliente</span></header>
              <div className="wrap"><table className="tbl">
                <thead><tr><th>Cliente</th><th className="r">Mesi</th><th className="r">Lead</th><th className="r">Demo svolte</th><th className="r">Chiusure</th><th className="r">Chiusura</th><th className="r">Contratto</th><th className="r">Incassato</th><th className="r">Costo/chiusura</th></tr></thead>
                <tbody>
                  {mapValuta(perCli, (r) => r.z.id, (r) => (
                    <tr key={r.key} className="clic" onClick={() => { cambiaAccount(r.z.id); setVista("vendite"); }} title="Apri il ciclo di vendita del cliente">
                      <td style={{ paddingLeft: 14 + r.liv * 18 }}>{nomeGruppo("vendite", r)}</td>
                      <td className="r">{int(r.mesi)}<div className="sub mono">fino a {meseIt(r.ultimo)}</div></td>
                      <td className="r">{int(r.lead)}</td>
                      <td className="r">{int(r.demoSvolte)}</td>
                      <td className="r">{int(r.chiusure)}</td>
                      <td className="r">{r.demoSvolte ? pctv((r.chiusure / r.demoSvolte) * 100) : "—"}</td>
                      <td className="r">{eur0(r.valoreContratto)}</td>
                      <td className="r">{eur0(r.incassato)}</td>
                      <td className="r">{r.chiusure && r.spesa ? eur(r.spesa / r.chiusure) : "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table></div>
              {!perCli.length && <div className="empty"><b>Nessun periodo caricato</b>I mesi si registrano dentro ogni cliente, a mano, dal registro dei venditori o da un file KPI.</div>}
            </section>
          );
        })()}

        {attivo && <section className="card">
          <header><h2>Periodi registrati</h2><span className="sp eyebrow">show up = svolte su fissate</span></header>          <div className="wrap">
            <table className="tbl">
              <thead><tr>
                <th>Mese</th>{!attivo && <th>Cliente</th>}
                <th className="r">Lead</th><th className="r">Chiamate</th>
                <th className="r">Discovery</th><th className="r">Show up</th>
                <th className="r">Demo</th><th className="r">Show up</th>
                <th className="r">Chiusure</th><th className="r">Chiusura</th>
                <th className="r">Contratto</th><th className="r">Incassato</th><th className="r">Costo/chiusura</th>
              </tr></thead>
              <tbody>
                {mapValuta(righe, (r) => r.clienteId, (r) => (
                  <tr key={r.id} className="clic" onClick={() => apri("vendite", data.vendite.find((x) => x.id === r.id))}>
                    <td>
                      <b>{meseIt(r.periodo)}</b>
                      {r.daGiornate && <div className="sub">{r.daGiornate} giornate caricate</div>}
                      {r.auto && <div className="sub">chiusure da schede</div>}
                    </td>
                    {!attivo && <td className="sub">{nomeCli(r.clienteId)}</td>}
                    <td className="r">{int(r.lead)}</td>
                    <td className="r">{int(r.chiamate)}</td>
                    <td className="r">{int(r.discoverySvolte)}<div className="sub mono">di {int(r.discoveryFissate)}</div></td>
                    <td className="r">{r.discoveryFissate ? pctv((r.discoverySvolte / r.discoveryFissate) * 100) : "—"}</td>
                    <td className="r">{int(r.demoSvolte)}<div className="sub mono">di {int(r.demoFissate)}</div></td>
                    <td className="r">{r.demoFissate ? pctv((r.demoSvolte / r.demoFissate) * 100) : "—"}</td>
                    <td className="r">{int(r.chiusure)}</td>
                    <td className="r">{r.demoSvolte ? pctv((r.chiusure / r.demoSvolte) * 100) : "—"}</td>
                    <td className="r">{eur0(r.valoreContratto)}</td>
                    <td className="r">{eur0(r.incassato)}</td>
                    <td className="r">{r.chiusure ? eur(r.spesa / r.chiusure) : "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {!righe.length && (
            <div className="empty"><b>Nessun periodo caricato</b>Aggiungi un mese con chiamate, discovery e demo: le chiusure arrivano dalle schede della sezione Chiusure.</div>
          )}
        </section>}
      </>
    ));
  };

  const SezioneTask = () => (
    <>
      <section className="card">
        <header>
          <div className="filtri">
            <button className="btn ghost mini" data-on={vistaTask === "bacheca" ? "1" : "0"} onClick={() => setVistaTask("bacheca")}>Bacheca</button>
            <button className="btn ghost mini" data-on={vistaTask === "elenco" ? "1" : "0"} onClick={() => setVistaTask("elenco")}>Elenco</button>
            <select value={filtri.area} onChange={(e) => setFiltri({ ...filtri, area: e.target.value })} aria-label="Filtra per area">
              <option value="">Tutte le aree</option>
              {data.aree.map((a) => <option key={a.id} value={a.id}>{a.nome}</option>)}
            </select>
            <select value={filtri.prio} onChange={(e) => setFiltri({ ...filtri, prio: e.target.value })} aria-label="Filtra per priorità">
              <option value="">Tutte le priorità</option>
              {PRIORITA.map((p) => <option key={p.v} value={p.v}>{p.l}</option>)}
            </select>
            {!attivo && (
              <select value={filtri.cli} onChange={(e) => setFiltri({ ...filtri, cli: e.target.value })} aria-label="Filtra per cliente">
                <option value="">Tutti i clienti</option>
                {attivi.map((z) => <option key={z.id} value={z.id}>{z.nome}</option>)}
              </select>
            )}
            <label className="sub riga" style={{ gap: 6 }}>
              <input type="checkbox" style={{ width: 15, height: 15, accentColor: "var(--acc)" }}
                checked={filtri.fatte} onChange={(e) => setFiltri({ ...filtri, fatte: e.target.checked })} />
              mostra fatte
            </label>
          </div>
          <span className="sp eyebrow">
            {taskBase.filter((t) => t.stato !== "fatto").length} aperti · {taskBase.filter((t) => t.stato !== "fatto" && t.scadenza && t.scadenza < oggi()).length} in ritardo
          </span>
        </header>
      </section>

      {vistaTask === "bacheca" ? (
        <div className="board">
          {STATI_TASK.map((s) => {
            const items = ordinaTask(taskBase.filter((t) => t.stato === s.v));
            const ore = items.reduce((a, b) => a + (b.stima || 0), 0);
            return (
              <div className="col" key={s.v}>
                <h3>{s.l}<span>{items.length}{ore ? ` · ${fmt(ore, 1)} h` : ""}</span></h3>
                <div className="list">
                  {items.map((t) => <CartaTask key={t.id} t={t} />)}
                  {!items.length && <div className="sub" style={{ padding: "8px 2px" }}>Vuota.</div>}
                </div>
              </div>
            );
          })}
        </div>
      ) : (
        <section className="card">
          {gruppiTask().map((g) => (
            <div key={g.t}>
              <div className="gruppo">{g.t} · {g.items.length}</div>
              {g.items.map((t) => <RigaTask key={t.id} t={t} />)}
            </div>
          ))}
          {!gruppiTask().length && <div className="empty"><b>Niente in elenco</b>Con questi filtri non c'è nessun task aperto.</div>}
        </section>
      )}
    </>
  );

  return (
    <div className="crm" data-tema={tema}>
      <style>{CSS}</style>
      <div className="shell">
        <nav className="rail" data-compatto={compatto ? "1" : "0"}>
          <div className="brand">
            <Logo src={logoAgenzia} nome={data.agenzia.nome} size={compatto ? 32 : 34} />
            {!compatto && (
              <>
                <div className="testo"><b>{data.agenzia.nome}</b><span>{data.agenzia.sottotitolo}</span></div>
                <button className="tgl" onClick={() => setCompatto(true)} title="Riduci il menu a icone" aria-label="Riduci il menu a icone">«</button>
              </>
            )}
          </div>
          {compatto && (
            <button className="tgl solo" onClick={() => setCompatto(false)} title="Espandi il menu" aria-label="Espandi il menu">»</button>
          )}
          <div className="switch">
            {compatto ? (
              <button className="accbtn" onClick={() => setCompatto(false)}
                title={attivo ? `Account: ${attivo.nome}` : "Tutti i clienti"} aria-label="Cambia account">
                {attivo ? <Logo src={attivo.logo} nome={attivo.nome} size={28} /> : "▣"}
              </button>
            ) : (
              (
                <>
                  <span className="eyebrow">Account</span>
                  <select value={cliAtt} onChange={(e) => cambiaAccount(e.target.value)} aria-label="Cambia account">
                    <option value="">▣ Tutti i clienti</option>
                    {ordinaGerarchia().map(({ z, liv }) => (
                      <option key={z.id} value={z.id}>{liv ? `${"\u00A0".repeat(liv * 4)}↳ ` : ""}{z.nome}</option>
                    ))}
                  </select>
                </>
              )
            )}
          </div>
          {NAV.map((n) => (
            <button key={n.id} className="navb" title={T(n.l)} data-on={vista === n.id ? "1" : "0"}
              data-badge={n.al ? "1" : "0"} onClick={() => { setVista(n.id); setQ(""); }}>
              <i><IcoNav id={n.id} /></i>
              <span className="etichetta">{T(n.l)}</span>
              {n.n !== undefined && <u data-al={n.al ? "1" : "0"}>{n.n}</u>}
            </button>
          ))}
          <div className="railfine">
            <div className="railfoot">
              <span>I dati sono salvati nel cloud, nel tuo spazio di lavoro.</span>
              <button className="navb" onClick={reset} style={{ padding: "6px 10px", fontSize: 12 }}>Riparti dai dati di esempio</button>
            </div>
          </div>
          {onEsci && (
            <div className="esci-fisso">
              <button className="navb" onClick={onEsci} title={utente ? `${T(t("nav.esci"))} · ${utente}` : T(t("nav.esci"))}>
                <i><IcoNav id="esci" /></i>
                <span className="etichetta">{T(t("nav.esci"))}</span>
              </button>
            </div>
          )}
        </nav>

        <div className="main">
          <header className="top">
            {vista !== "cruscotto" && (
              <button className="btn ghost mini" onClick={() => setVista("cruscotto")} title="Torna al cruscotto">‹</button>
            )}
            {attivo && <Logo src={attivo.logo} nome={attivo.nome} size={34} />}
            <div>
              <h1>{titoli[vista] ? T(titoli[vista][0]) : ""}</h1>
              <p>
                {attivo && attivo.padreId ? `${nomeCli(attivo.padreId)} › ` : ""}
                {attivo && vista !== "cruscotto" ? `${attivo.nome} · ` : ""}
                {titoli[vista] ? T(titoli[vista][1]) : ""}
              </p>
            </div>
            <span style={{ marginLeft: "auto" }} />
            {attivo && haFigli && (
              <button className="btn ghost mini" data-on={consolidato ? "1" : "0"} onClick={() => setConsolidato(!consolidato)}
                title="Includere o escludere i dati dei sottoclienti">
                {consolidato ? `Consolidato · ${figli(attivo.id).length + 1} account` : "Solo questo account"}
              </button>
            )}
            {salvato && <span className="saved">salvato</span>}
            {(erroreArchivio || erroreSalvataggio) && (
              <span role="alert" title={erroreArchivio || erroreSalvataggio}
                style={{ color: "#fff", background: "#B42318", padding: "5px 10px", borderRadius: 6, maxWidth: 380, whiteSpace: "normal", lineHeight: 1.35, fontSize: 12 }}>
                {erroreArchivio || erroreSalvataggio}
              </span>
            )}
            {["clienti", "contatti", "campagne", "advertising", "file"].includes(vista) && (
              <input placeholder="Cerca" value={q} onChange={(e) => setQ(e.target.value)} style={{ maxWidth: 180 }} />
            )}
            {vista === "scheda" && <button className="btn" onClick={() => apri("clienti", attivo)}>Modifica scheda</button>}
            {bottoneAggiungi && <button className="btn" onClick={() => apri(bottoneAggiungi)}>+ {t("a.aggiungi")}</button>}

            {attivo && (
              <div className="lingua valuta">
                <label htmlFor="sel-valuta">{t("a.valuta")}</label>
                <select id="sel-valuta" value={attivo.valuta || ""} aria-label={t("a.valuta")}
                  onChange={(e) => upd("clienti", (l) => l.map((z) => (z.id === attivo.id ? { ...z, valuta: e.target.value } : z)))}>
                  <option value="">{data.agenzia.valuta || "EUR"} · {t("a.predefinita")}</option>
                  {VALUTE.map((c) => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
            )}

            <div className="lingua">
              <label htmlFor="sel-lingua">{t("a.lingua")}</label>
              <select id="sel-lingua" value={lingua} aria-label={t("a.lingua")}
                onChange={(e) => {
                  const nuova = e.target.value;
                  const L = LINGUE.find((x) => x.v === nuova) || LINGUE[0];
                  const attuale = data.agenzia.valuta || "EUR";
                  upd("agenzia", () => ({ ...data.agenzia, lingua: nuova }));
                  /* La valuta non la cambio di nascosto: gli importi già inseriti
                     non si convertono, cambierebbe solo il simbolo davanti. */
                  if (L.val !== attuale) chiedi({
                    titolo: `Cambiare anche la valuta in ${L.val}?`,
                    testo: `${L.l} di solito si accompagna a ${L.val}, mentre ora l'agenzia usa ${attuale}.\n\nAttenzione: l'app non converte gli importi, cambia il simbolo con cui li mostra. Accetta solo se i numeri che hai inserito sono già in ${L.val}. I clienti con una valuta propria restano come sono.`,
                    etichetta: `Usa ${L.val}`,
                    azione: () => upd("agenzia", () => ({ ...data.agenzia, lingua: nuova, valuta: L.val })),
                    etichetta2: `Tieni ${attuale}`,
                    azione2: () => {},
                  });
                }}>
                {LINGUE.map((x) => <option key={x.v} value={x.v}>{x.l}</option>)}
              </select>
            </div>
          </header>

          <div className="body">
            {/* ================= AGENZIA ================= */}
            {!attivo && vista === "cruscotto" && (() => {
              /* La panoramica parla solo dei clienti principali: i sottoclienti si guardano
                 entrando nel loro madre, dove l'interruttore consolidato li somma. */
              const principali = attivi.filter((z) => !z.padreId);
              const redd = redditivitaClienti(data);
              const idsPrinc = principali.map((z) => z.id);
              /* Valute diverse nello stesso totale: la somma non corrisponde a nulla */
              const valuteUsate = [...new Set(attivi.map((z) => valutaCliente(data, z.id)))];
              const mesi = righeMensili(idsPrinc);
              const ora = mesi[mesi.length - 1] || { spesa: 0, lead: 0, chiusure: 0, contratto: 0, incassato: 0, cpl: 0, ritorno: 0 };
              const prima = mesi[mesi.length - 2] || null;
              /* Chiusure e contratti arrivano dopo la spesa: per questi riquadri vale l'ultimo mese che li ha */
              const iV = mesi.map((x) => !!(x.chiusure || x.contratto)).lastIndexOf(true);
              const oraV = iV > -1 ? mesi[iV] : ora, primaV = iV > 0 ? mesi[iV - 1] : null;
              const meseV = iV > -1 && oraV.mm !== ora.mm ? ` · ${meseIt(oraV.mm)}` : "";
              const deltaV = (a, b) => {
                if (!primaV || !b) return null;
                const v = ((a - b) / b) * 100;
                if (!isFinite(v) || Math.abs(v) < 0.5) return { v: 0, buono: true };
                return { v, buono: v > 0 };
              };
              const delta = (a, b, meglioGiu) => {
                if (!prima || !b) return null;
                const v = ((a - b) / b) * 100;
                if (!isFinite(v) || Math.abs(v) < 0.5) return { v: 0, buono: true };
                return { v, buono: meglioGiu ? v < 0 : v > 0 };
              };
              const Delta = ({ d }) => !d ? <span className="sub">nessun mese precedente</span> : (
                <span className={d.v === 0 ? "sub" : d.buono ? "pos" : "neg"}>
                  {d.v === 0 ? "stabile" : `${d.v > 0 ? "▲" : "▼"} ${fmt(Math.abs(d.v), 0)}%`}
                  <span className="sub" style={{ marginLeft: 6 }}>sul mese prima</span>
                </span>
              );

              /* Distribuzione della spesa: per piattaforma, dalle campagne attive */
              const perPiatt = (() => {
                const mm = new Map();
                m.attive.filter((k) => idsPrinc.includes(k.clienteId))
                  .forEach((k) => mm.set(k.piattaforma, (mm.get(k.piattaforma) || 0) + k.d.sp));
                const l = [...mm.entries()].map(([k, v]) => ({ k: et(PIATTAFORME, k), v })).sort((a, b) => b.v - a.v);
                return l.length > 5 ? [...l.slice(0, 4), { k: "Altro", v: l.slice(4).reduce((s, x) => s + x.v, 0) }] : l;
              })();
              const spesaPiatt = perPiatt.reduce((s, x) => s + x.v, 0);

              /* Una riga per cliente, con lo stato che riassume il giudizio */
              const rigaPerf = (p) => {
                const r = righeMensili([p.id]);
                const u = r[r.length - 1] || { lead: 0, cpl: 0, demo: 0, chiusure: 0, contratto: 0, resa: 0, tentativi: 0, suChiamate: false };
                /* I numeri commerciali arrivano spesso dopo la spesa (il mese in corso ha già Meta ma
                   non ancora il ciclo di vendita): per demo, chiusure, contratti e resa prendo l'ultimo
                   mese che li ha, e lo dichiaro sotto il numero se non è lo stesso della spesa. */
                const uv = [...r].reverse().find((x) => x.demo || x.chiusure || x.contratto || x.incassato);
                const vend = uv ? { demo: uv.demo, chiusure: uv.chiusure, contratto: uv.contratto, resa: uv.resa, tentativi: uv.tentativi,
                  suChiamate: uv.suChiamate, meseVend: uv.mm !== u.mm ? uv.mm : "" } : { meseVend: "" };
                /* Gli avvisi comprendono i sottoclienti: un problema critico su un loro
                   account non deve restare invisibile solo perché i numeri qui sono del madre. */
                const suoi = gruppo(p.id);
                /* Lo stato parla dei risultati: i task in ritardo restano negli Avvisi ma non lo decidono,
                   e le campagne ferme (pagamento, account) hanno uno stato loro, non "critico" */
                const delCli = m.avvisi.filter((a) => suoi.includes(a.clienteId) && a.metrica !== "Task");
                const fermi = delCli.filter((a) => a.gravita === "critica" && ["Pagamento", "Account"].includes(a.metrica)).length;
                const critici = delCli.filter((a) => a.gravita === "critica" && !["Pagamento", "Account"].includes(a.metrica)).length;
                const attenz = delCli.filter((a) => a.gravita === "attenzione").length;
                const daFigli = delCli.filter((a) => a.clienteId !== p.id && a.gravita !== "nota").length;
                const red = redd[p.id] && redd[p.id].attivo ? redd[p.id] : null;
                const redNo = redd[p.id] && !redd[p.id].attivo ? redd[p.id] : null;
                const stato = critici ? "critico" : fermi ? "fermo" : attenz ? (red ? "buono" : "attenzione") : p.modo === "nd" ? (red ? "ottimo" : "nd") : p.perf >= 1 || red ? "ottimo" : p.perf >= 0.8 ? "buono" : "attenzione";
                return { ...p, ...u, ...vend, stato, critici, attenz, daFigli, fermi, red, redNo };
              };
              const perf = m.portafoglio.filter((p) => idsPrinc.includes(p.id)).map(rigaPerf).sort((a, b) => b.sp - a.sp);
              /* Con la freccia, sotto il cliente principale compaiono i suoi sottoclienti con i loro numeri */
              const perfVisibili = perf.flatMap((p) => {
                const sotto = ordinaGerarchia(p.id, 1);
                const r = { ...p, liv: 0, nSotto: sotto.length };
                if (!sotto.length || !apertoCli("perf", p.id)) return [r];
                return [r, ...sotto.map((x) => m.portafoglio.find((q) => q.id === x.z.id)).filter(Boolean).map((q) => ({ ...rigaPerf(q), liv: 1, nSotto: 0 }))];
              });
              const colore = { ottimo: "var(--ok)", buono: "var(--ok)", attenzione: "var(--warn)", critico: "var(--acc-t)", fermo: "var(--acc-t)", nd: "var(--ink3)" };
              const testoStato = { ottimo: "Ottimo", buono: "Buono", attenzione: "Attenzione", critico: "Critico", fermo: "Campagne ferme", nd: "Senza obiettivo" };

              return conValuta(valutaComune(idsPrinc), () => (
                <>
                  {meta && metaInfo && !metaFb && (
                    <section className="card">
                      <div className="riga" style={{ gap: 12, padding: "14px 18px", flexWrap: "wrap" }}>
                        <span aria-hidden="true" style={{ width: 34, height: 34, borderRadius: 8, background: "#1877F2", color: "#fff", display: "inline-flex", alignItems: "center", justifyContent: "center", fontWeight: 700 }}>M</span>
                        <div style={{ flex: 1, minWidth: 220 }}><b>Collega il tuo profilo Meta</b><div className="sub">Il CRM legge in sola lettura campagne, spesa e risultati dei tuoi account pubblicitari e li aggiorna ogni mattina.</div></div>
                        <button className="btn" disabled={!!metaStato.attesa} onClick={collegaFacebook}>{metaStato.attesa === "facebook" ? "Apro Meta…" : "Collega a Meta"}</button>
                        <button className="btn ghost mini" onClick={() => setVista("impostazioni")}>Altre piattaforme</button>
                      </div>
                    </section>
                  )}
                  {valuteUsate.length > 1 && (
                    <article className="al" data-g="attenzione">
                      <div className="meta">
                        <span className="tag" data-k="late">{valuteUsate.join(" + ")}</span>
                        <span className="misura">i totali qui sotto le sommano</span>
                      </div>
                      <h3>I clienti usano valute diverse</h3>
                      <dl>
                        <dt>Cosa comporta</dt>
                        <dd>Le somme a livello di agenzia mettono insieme {valuteUsate.join(", ")} come se fossero la stessa cosa: il numero che ne esce non corrisponde a niente. L'app mostra ogni cliente nella sua valuta, ma non converte — servirebbe un tasso di cambio, e per farlo bene quello del giorno di ogni singola spesa.</dd>
                        <dt>Cosa fare</dt>
                        <dd>Leggi gli importi entrando nel singolo cliente, dove la valuta è una sola. Qui a livello di agenzia restano attendibili i conteggi che non dipendono dalla valuta: campagne, lead, demo, chiusure.</dd>
                      </dl>
                    </article>
                  )}

                  <div className="kpis">
                    <div className="kpi"><IcoKpi t="Spesa gestita" /><div className="eyebrow">Spesa del mese</div>
                      <div className="v">{eur0(ora.spesa)}</div>
                      <div className="s"><Delta d={delta(ora.spesa, prima && prima.spesa)} /></div></div>
                    <div className="kpi"><IcoKpi t="Lead" /><div className="eyebrow">Lead generati</div>
                      <div className="v">{int(ora.lead)}</div>
                      <div className="s"><Delta d={delta(ora.lead, prima && prima.lead)} /></div></div>
                    <div className="kpi"><IcoKpi t="CPL" /><div className="eyebrow">CPL medio</div>
                      <div className="v">{ora.cpl ? eur(ora.cpl) : "—"}</div>
                      <div className="s"><Delta d={delta(ora.cpl, prima && prima.cpl, true)} /></div></div>
                    <div className="kpi"><IcoKpi t="Costo per chiusura" /><div className="eyebrow">Costo per cliente</div>
                      <div className="v">{oraV.chiusure && oraV.spesa ? eur0(oraV.spesa / oraV.chiusure) : "—"}</div>
                      <div className="s">{int(oraV.chiusure)} {oraV.chiusure === 1 ? "chiusura" : "chiusure"} {meseV ? `a ${meseIt(oraV.mm)}` : "nel mese"}</div></div>
                    <div className="kpi"><IcoKpi t="Valore contratti" /><div className="eyebrow">Contratti firmati{meseV}</div>
                      <div className="v">{eur0(oraV.contratto)}</div>
                      <div className="s"><Delta d={deltaV(oraV.contratto, primaV && primaV.contratto)} /></div></div>
                    <div className="kpi"><IcoKpi t="Ritorno" /><div className="eyebrow">Ritorno sulla spesa{meseV}</div>
                      <div className="v">{oraV.ritorno ? `${fmt(oraV.ritorno)}×` : "—"}</div>
                      <div className="s">
                        {m.critici ? <span className="neg">{m.critici} avvisi critici</span> : `${m.avvisi.length} avvisi aperti`}
                        <span className="sub" style={{ marginLeft: 6 }}>su tutti gli account</span>
                      </div></div>
                  </div>

                  <div className="two">
                    <section className="card">
                      <header>
                        <h2>Andamento del portafoglio</h2>
                        <span className="sp eyebrow">contratti e spesa, mese per mese</span>
                        <span className="riga" style={{ gap: 12, fontSize: 11.5 }}>
                          <span className="riga" style={{ gap: 6 }}><span style={{ width: 14, height: 2, background: "#D8455F" }} />Contratti</span>
                          <span className="riga" style={{ gap: 6 }}><span style={{ width: 14, height: 2, background: "#8A8A96" }} />Spesa</span>
                        </span>
                      </header>
                      <div style={{ padding: "10px 12px 4px" }}><Andamento righe={mesi.slice(-12)} /></div>
                    </section>

                    <section className="card">
                      <header><h2>Distribuzione della spesa</h2><span className="sp eyebrow">per piattaforma</span></header>
                      {spesaPiatt ? <Anello voci={perPiatt} totale={spesaPiatt} etichetta="Attive" />
                        : <div className="empty">Nessuna campagna attiva con spesa registrata.</div>}
                    </section>
                  </div>

                  <section className="card">
                    <header>
                      <h2>Performance dei clienti</h2>
                      <span className="sp eyebrow">clienti principali · ▸ apre i sottoclienti · ultimo mese con dati, per spesa e per vendite</span>
                      <button className="btn ghost mini" onClick={() => setVista("mesi")}>Riepilogo mensile</button>
                    </header>
                    <div className="wrap">
                      <table className="tbl">
                        <thead><tr>
                          <th>Cliente</th><th className="r">Spesa</th><th className="r">Lead</th><th className="r">CPL</th>
                          <th className="r">Demo</th><th className="r">Chiusure</th><th className="r">Contratti</th>
                          <th className="r">Sforzo/Resa</th><th className="r">Resa ads</th><th>Stato</th>
                        </tr></thead>
                        <tbody>
                          {mapValuta(perfVisibili, (p) => p.id, (p) => (
                            <tr key={p.id} className="clic" onClick={() => cambiaAccount(p.id)}>
                              <td style={{ paddingLeft: 14 + p.liv * 18 }}><div className="riga">
                                {p.liv === 0 && frecciaCli("perf", p.id, p.nSotto)}
                                {p.liv > 0 && <span className="sub">↳</span>}
                                <Logo src={p.logo} nome={p.nome} size={p.liv ? 22 : 26} /><div>
                                <b>{p.nome}</b>
                                <div className="sub">
                                  {p.nCamp} campagne
                                  {p.liv === 0 && p.nSotto > 0 && ` · ${p.nSotto} sottoclienti esclusi dai numeri`}
                                </div></div></div></td>
                              <td className="r">{p.sp ? eur0(p.sp) : "—"}</td>
                              <td className="r">{p.lead ? int(p.lead) : "—"}</td>
                              <td className="r">{p.cpl ? eur(p.cpl) : "—"}</td>
                              <td className="r">{p.demo ? int(p.demo) : "—"}{p.meseVend && p.demo ? <div className="sub">{meseIt(p.meseVend)}</div> : null}</td>
                              <td className="r">{p.chiusure ? int(p.chiusure) : "—"}{p.meseVend && p.chiusure ? <div className="sub">{meseIt(p.meseVend)}</div> : null}</td>
                              <td className="r">{p.contratto ? eur0(p.contratto) : "—"}{p.meseVend && p.contratto ? <div className="sub">{meseIt(p.meseVend)}</div> : null}</td>
                              <td className={"r " + (p.resa ? (p.cpl && p.resa >= p.cpl ? "pos" : p.cpl ? "neg" : "") : "")}
                                title={p.resa ? `${eur(p.resa)} incassati per ogni ${p.suChiamate ? "chiamata fatta" : "contatto entrato nel funnel"} (${int(p.tentativi)}). ${p.cpl ? `Il contatto costa ${eur(p.cpl)}: ${p.resa >= p.cpl ? "il funnel è in profitto" : "ogni contatto costa più di quanto rende"}.` : ""}` : "Serve il numero di chiamate o di lead del mese"}>
                                {p.resa ? eurDi(p.resa, p.valuta) : "—"}
                                {p.resa > 0 && <div className="sub">{p.suChiamate ? "per chiamata" : "per contatto"}{p.meseVend ? ` · ${meseIt(p.meseVend)}` : ""}</div>}
                              </td>
                              <td className="r">{p.testo}</td>
                              <td>
                                <span className="riga" style={{ gap: 7 }}>
                                  <span style={{ width: 7, height: 7, borderRadius: "50%", background: colore[p.stato], flex: "none" }} />
                                  {testoStato[p.stato]}
                                  {p.daFigli > 0 && <span className="sub">({p.daFigli} dai sottoclienti)</span>}
                                </span>
                                {p.stato === "fermo" && <div className="sub">{p.fermi} {p.fermi === 1 ? "campagna ferma" : "campagne ferme"} per pagamento o account</div>}
                                {p.red && <div className="sub pos">in attivo · contratti {fmt(p.red.ritorno)}× la spesa</div>}
                                {p.redNo && <div className="sub neg">contratti {fmt(p.redNo.ritorno)}× la spesa{p.redNo.soglia > 1 ? `, pareggio a ${fmt(p.redNo.soglia)}×` : ""}</div>}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    {!perf.length && <div className="empty"><b>Nessun cliente attivo</b>Aggiungi la prima azienda che segui.</div>}
                  </section>

                  <div className="two">
                    <section className="card">
                      <header><h2>Avvisi da gestire</h2><span className="sp" /><button className="btn ghost mini" onClick={() => setVista("avvisi")}>Vedi tutti</button></header>
                    <div style={{ padding: 14, display: "flex", flexDirection: "column", gap: 12 }}>
                      {m.avvisi.slice(0, 3).map((a, i) => <Avviso key={i} a={a} dove={`${nomeCli(a.clienteId)}${a.campagna ? ` · ${a.campagna}` : ""}`} />)}
                      {!m.avvisi.length && <div className="empty"><b>Tutto entro le soglie</b>Nessuna campagna fuori benchmark in questo momento.</div>}
                    </div>
                  </section>
                  <section className="card">
                    <header><h2>Task in scadenza</h2><span className="sp" /><button className="btn ghost mini" onClick={() => setVista("task")}>Apri</button></header>
                    <div className="acts">
                      {ordinaTask(data.task.filter((t) => t.stato !== "fatto")).slice(0, 7).map((t) => <RigaTask key={t.id} t={t} />)}
                      {!data.task.some((t) => t.stato !== "fatto") && <div className="empty"><b>Niente in agenda</b>Nessun task aperto.</div>}
                    </div>
                  </section>
                </div>
                </>
              ));
            })()}

            {!attivo && vista === "clienti" && (
              <>
              <section className="card">
                <div className="wrap">
                  <table className="tbl">
                    <thead><tr><th>Cliente</th><th>Stato</th><th className="r">Compenso</th><th className="r">Campagne</th><th className="r">Speso</th><th className="r">Lead</th><th className="r">CPL</th><th className="r">Sforzo/Resa</th><th className="r">Resa ads</th><th className="r">Task</th><th className="r">Avvisi</th><th /></tr></thead>
                    <tbody>
                      {mapValuta(q
                        /* Cercando si vedono il cliente trovato, il suo principale e i suoi sottoclienti, senza tendina */
                        ? ordinaGerarchia().filter(({ z }) => [z.id, ...discendenti(z.id), z.padreId].some((id) => { const y = cliente(id); return y && cerca(y.nome + " " + y.settore); }))
                        : alberoCli("clienti", ordinaGerarchia()), (x) => x.z.id, ({ z, liv, nSotto = 0 }) => {
                        const p = m.portafoglio.find((x) => x.id === z.id) || { sp: 0, nCamp: 0, nAvvisi: 0, testo: "—", modo: "nd", perf: 0 };
                        const nt = data.task.filter((t) => t.clienteId === z.id && t.stato !== "fatto").length;
                        const nf = figli(z.id).length;
                        const rz = righeMensili([z.id]);
                        const uz = rz[rz.length - 1] || { lead: 0, cpl: 0, resa: 0, tentativi: 0, suChiamate: false };
                        return (
                          <tr key={z.id} className="clic" onClick={() => cambiaAccount(z.id)}>
                            <td style={{ paddingLeft: 14 + liv * 26 }}>
                              <div className="riga">
                                {frecciaCli("clienti", z.id, nSotto)}
                                {liv > 0 && <span className="sub" style={{ marginLeft: -14 }}>↳</span>}
                                <Logo src={z.logo} nome={z.nome} size={liv ? 24 : 30} />
                                <div><b>{z.nome}</b><div className="sub">
                                  {[z.settore, z.citta].filter(Boolean).join(" · ") || "da completare"}
                                  {nf > 0 && <span className="pill" style={{ marginLeft: 6 }}>{nf} sottoclienti</span>}
                                  {z.valuta && z.valuta !== (data.agenzia.valuta || "EUR") && (
                                    <span className="pill" style={{ marginLeft: 6 }}>{z.valuta}</span>
                                  )}
                                </div></div>
                              </div>
                            </td>
                            <td><span className="tag" data-k={z.stato}>{et(STATI_CLIENTE, z.stato)}</span></td>
                            <td className="r">{z.fee ? `${eurDi(z.fee, valutaCompenso(data, z.id))}/mese` : "—"}</td>
                            <td className="r">{p.nCamp}</td>
                            <td className="r">{p.sp ? eurDi(p.sp, z.valuta) : "—"}</td>
                            <td className="r">{uz.lead ? int(uz.lead) : "—"}</td>
                            <td className="r">{uz.cpl ? eurDi(uz.cpl, z.valuta) : "—"}</td>
                            <td className={"r " + (uz.resa ? (uz.cpl && uz.resa >= uz.cpl ? "pos" : uz.cpl ? "neg" : "") : "")}
                              title={uz.resa ? `${eur(uz.resa)} incassati per ogni ${uz.suChiamate ? "chiamata fatta" : "contatto entrato nel funnel"} (${int(uz.tentativi)})` : "Serve il numero di chiamate o di lead del mese"}>
                              {uz.resa ? eurDi(uz.resa, z.valuta) : "—"}
                              {uz.resa > 0 && <div className="sub">{uz.suChiamate ? "per chiamata" : "per contatto"}</div>}
                            </td>
                            <td className={"r " + (p.modo !== "nd" && p.perf < 1 ? "neg" : "")}>{p.testo}</td>
                            <td className="r">{nt || "—"}</td>
                            <td className="r">{p.nAvvisi ? <span className="tag" data-k="late">{p.nAvvisi}</span> : <span className="sub">ok</span>}</td>
                            <td className="r">
                              <div className="riga" style={{ justifyContent: "flex-end", gap: 6 }}>
                                {!z.padreId && (
                                  <button className="btn ghost mini" onClick={(e) => { e.stopPropagation(); apri("clienti", { padreId: z.id }, true); }}>
                                    + Sottocliente
                                  </button>
                                )}
                                <button className="btn ghost mini" title="Mettilo da parte senza perdere lo storico"
                                  onClick={(e) => { e.stopPropagation(); archiviaCliente(z); }}>Archivia</button>
                                <button className="btn ghost mini" title="Cancellalo con tutti i suoi dati"
                                  onClick={(e) => { e.stopPropagation(); eliminaCliente(z); }}>Elimina</button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
                {!attivi.length && <div className="empty"><b>Nessun cliente</b>Aggiungi la prima azienda per cui compri media.</div>}
              </section>

                {data.clienti.some((z) => z.archiviato === "si") && (
                  <section className="card">
                    <header>
                      <h2>Clienti archiviati</h2>
                      <span className="sp eyebrow">{data.clienti.filter((z) => z.archiviato === "si").length} fuori dai conti</span>
                      <button className="btn ghost mini" data-on={mostraArchCli ? "1" : "0"} onClick={() => setMostraArchCli(!mostraArchCli)}>
                        {mostraArchCli ? "Nascondi" : "Mostra"}
                      </button>
                    </header>
                    {!mostraArchCli ? (
                      <div className="nota" style={{ paddingTop: 14 }}>
                        Non compaiono nell'elenco, nei totali, nel semaforo né negli avvisi. Il loro storico è conservato:
                        premi Mostra per rivederli e riattivarli.
                      </div>
                    ) : (
                      <div className="wrap">
                        <table className="tbl">
                          <thead><tr>
                            <th>Cliente</th><th>Archiviato il</th><th className="r">Voci in deposito</th>
                            <th className="r">Ultima spesa</th><th className="r">Contratti</th><th />
                          </tr></thead>
                          <tbody>
                            {mapValuta(data.clienti.filter((z) => z.archiviato === "si"), (z) => z.id, (z) => {
                              const voci = data.deposito.filter((v2) => v2.clienteId === z.id || (v2.ids || []).includes(z.id));
                              const suoi = (c) => voci.flatMap((v2) => (v2.dati[c] || [])).filter((x) => x.clienteId === z.id);
                              const n = COLL_CLI.reduce((s2, c) => s2 + suoi(c).length, 0);
                              const sp = suoi("campagne").reduce((s2, k) => s2 + (k.spesa || 0), 0)
                                + suoi("periodi").reduce((s2, pz) => s2 + (pz.spesa || 0), 0);
                              const ct = suoi("chiusure").reduce((s2, c) => s2 + parseNumIntl(c.valoreContratto), 0);
                              return (
                                <tr key={z.id}>
                                  <td>
                                    <div className="riga">
                                      <Logo src={z.logo} nome={z.nome} size={26} />
                                      <div>
                                        <b>{z.nome}</b>
                                        <div className="sub">{[z.settore, z.padreId ? `sottocliente di ${nomeCli(z.padreId)}` : ""].filter(Boolean).join(" · ") || "—"}</div>
                                      </div>
                                    </div>
                                  </td>
                                  <td className="sub">{z.archiviatoIl ? dataItA(z.archiviatoIl) : "—"}</td>
                                  <td className="r">{n ? int(n) : <span className="sub">nessuna</span>}</td>
                                  <td className="r">{sp ? eur0(sp) : "—"}</td>
                                  <td className="r">{ct ? eur0(ct) : "—"}</td>
                                  <td className="r">
                                    <button className="btn ghost mini" onClick={() => riattivaCliente(z)}>Riattiva</button>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}
                    {mostraArchCli && (
                      <div className="nota">
                        Riattivando puoi scegliere se riportare dentro tutto lo storico o ripartire da zero. Nel secondo caso i dati
                        restano qui e li reinserisci quando vuoi dalla scheda del cliente.
                      </div>
                    )}
                  </section>
                )}
              </>
            )}

            {!attivo && vista === "sottoclienti" && (() => {
              const madri = attivi.filter((z) => !z.padreId && figli(z.id).length > 0);
              if (!madri.length) return (
                <section className="card">
                  <header><h2>Sottoclienti</h2></header>
                  <div className="empty">
                    <b>Nessun sottocliente</b>
                    Servono quando gestisci un'azienda e insieme i suoi clienti finali. Li crei dalla scheda Sottoclienti nel
                    cruscotto di un cliente, dal pulsante in fondo alla sua riga nell'elenco Clienti, oppure importando un elenco.
                  </div>
                </section>
              );

              const tutti = madri.flatMap((z) => figli(z.id));
              const spesaTot = tutti.reduce((s2, z) => {
                const p = m.portafoglio.find((x) => x.id === z.id);
                return s2 + (p ? p.sp : 0);
              }, 0);
              const avvisiTot = m.avvisi.filter((a) => tutti.some((z) => z.id === a.clienteId)).length;

              return conValuta(valutaComune(tutti.map((z) => z.id)), () => (
                <>
                  <div className="kpis">
                    <div className="kpi"><IcoKpi t="Clienti" /><div className="eyebrow">Sottoclienti attivi</div>
                      <div className="v">{int(tutti.length)}</div>
                      <div className="s">sotto {madri.length} {madri.length === 1 ? "cliente" : "clienti"} principali</div></div>
                    <div className="kpi"><IcoKpi t="Spesa gestita" /><div className="eyebrow">Spesa complessiva</div>
                      <div className="v">{eur0(spesaTot)}</div>
                      <div className="s">esclusa dai totali dell'agenzia</div></div>
                    <div className="kpi"><IcoKpi t="Avvisi" /><div className="eyebrow">Avvisi aperti</div>
                      <div className="v">{int(avvisiTot)}</div>
                      <div className="s">{m.avvisi.filter((a) => tutti.some((z) => z.id === a.clienteId) && a.gravita === "critica").length} critici</div></div>
                    <div className="kpi"><IcoKpi t="Campagne" /><div className="eyebrow">Campagne</div>
                      <div className="v">{int(m.campagne.filter((k) => tutti.some((z) => z.id === k.clienteId)).length)}</div>
                      <div className="s">su tutti i sottoclienti</div></div>
                  </div>

                  {mapValuta(madri, (z) => z.id, (z) => {
                    const suoi = figli(z.id);
                    return (
                      <section className="card" key={z.id}>
                        <header className="clic" onClick={() => setApertiCli((x) => ({ ...x, ["sotto:" + z.id]: !x["sotto:" + z.id] }))}>
                          {frecciaCli("sotto", z.id, suoi.length)}
                          <h2>{z.nome}</h2>
                          <span className="sp eyebrow">{suoi.length} {suoi.length === 1 ? "sottocliente" : "sottoclienti"}{apertoCli("sotto", z.id) ? " · in ordine di cognome" : " · ▸ per vederli"}</span>
                          <button className="btn ghost mini" onClick={(e) => { e.stopPropagation(); cambiaAccount(z.id); }}>Apri il cliente</button>
                          <button className="btn ghost mini" onClick={(e) => { e.stopPropagation(); apri("clienti", { padreId: z.id }, true); }}>+ Sottocliente</button>
                        </header>
                        {apertoCli("sotto", z.id) && <div className="wrap">
                          <table className="tbl">
                            <thead><tr>
                              <th>Account</th><th>Stato</th><th className="r">Campagne</th><th>Data partenza</th><th className="r">Speso</th>
                              <th className="r">Lead</th><th className="r">CPL</th><th className="r">Resa</th><th className="r">Avvisi</th><th />
                            </tr></thead>
                            <tbody>
                              {mapValuta(suoi, (f) => f.id, (f) => {
                                const p = m.portafoglio.find((x) => x.id === f.id) || { sp: 0, nCamp: 0, nAvvisi: 0, testo: "—", modo: "nd", perf: 0 };
                                const r = righeMensili([f.id]);
                                const u = r[r.length - 1] || { lead: 0, cpl: 0 };
                                return (
                                  <tr key={f.id} className="clic" onClick={() => cambiaAccount(f.id)}>
                                    <td><div className="riga"><Logo src={f.logo} nome={f.nome} size={26} /><div>
                                      <b>{f.nome}</b><div className="sub">{[f.settore, f.citta].filter(Boolean).join(" · ") || "da completare"}</div>
                                    </div></div></td>
                                    <td><span className="tag" data-k={f.stato}>{et(STATI_CLIENTE, f.stato)}</span></td>
                                    <td className="r">{p.nCamp}</td>
                                    <td>
                                      {(() => {
                                        const pa = partenza(f.id);
                                        if (!pa.data) return <span className="sub">—</span>;
                                        return (
                                          <>
                                            {dataItA(pa.data)}
                                            <div className="sub">{pa.daCampagne ? `${int(giorniTra(pa.data, oggi()))} giorni di adv` : "in gestione, senza campagne"}</div>
                                          </>
                                        );
                                      })()}
                                    </td>
                                    <td className="r">{p.sp ? eurDi(p.sp, f.valuta || (attivi.find((z2) => z2.id === f.padreId) || {}).valuta) : "—"}</td>
                                    <td className="r">{u.lead ? int(u.lead) : "—"}</td>
                                    <td className="r">{u.cpl ? eurDi(u.cpl, f.valuta || (attivi.find((z2) => z2.id === f.padreId) || {}).valuta) : "—"}</td>
                                    <td className={"r " + (p.modo !== "nd" && p.perf < 1 ? "neg" : "")}>{p.testo}</td>
                                    <td className="r">{p.nAvvisi ? <span className="tag" data-k="late">{p.nAvvisi}</span> : <span className="sub">ok</span>}</td>
                                    <td className="r">
                                      <div className="riga" style={{ justifyContent: "flex-end", gap: 6 }}>
                                        <button className="btn ghost mini" title="Prepara la richiesta dei dati aggiornati"
                                          onClick={(e) => { e.stopPropagation(); setRichiesta({ nome: f.nome, testo: componiRichiesta(f) }); }}>Dati</button>
                                        <button className="btn ghost mini" title="Mettilo da parte senza perdere lo storico"
                                          onClick={(e) => { e.stopPropagation(); archiviaCliente(f); }}>Archivia</button>
                                        <button className="btn ghost mini" title="Cancellalo con tutti i suoi dati"
                                          onClick={(e) => { e.stopPropagation(); eliminaCliente(f); }}>Elimina</button>
                                      </div>
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>}
                      </section>
                    );
                  })}
                </>
              ));
            })()}

            {!attivo && vista === "avvisi" && (
              <div className="alerts">
                <button className="btn ghost mini" style={{ alignSelf: "flex-start" }} onClick={() => setVista("cruscotto")}>‹ Torna alla panoramica</button>
                {[...m.avvisi].sort((a, b) => (a.gravita === b.gravita ? 0 : a.gravita === "critica" ? -1 : 1))
                  .map((a, i) => <Avviso key={i} a={a} dove={`${nomeCli(a.clienteId)}${a.campagna ? ` · ${a.campagna}` : ""}`} />)}
                {!m.avvisi.length && <section className="card"><div className="empty"><b>Nessuna soglia superata</b>Quando una campagna esce dai benchmark la trovi qui.</div></section>}
              </div>
            )}

            {!attivo && vista === "impostazioni" && (
              <>
                {PannelloAspetto()}
                {PannelloPiattaforme()}
                {PannelloOrari()}
                {PannelloMetaAgenzia()}
                {PannelloTelegram()}
                {PannelloIntegrazioni()}
                {PannelloConnettore()}
                <section className="card">
                  <header><h2>Identità dell'agenzia</h2><span className="sp" /><button className="btn ghost mini" onClick={() => apri("agenzia")}>Modifica</button></header>
                  <div style={{ padding: 18, display: "flex", alignItems: "center", gap: 16 }}>
                    <Logo src={logoAgenzia} nome={data.agenzia.nome} size={64} />
                    <div>
                      <div style={{ fontSize: 18, fontWeight: 650 }}>{data.agenzia.nome}</div>
                      <div className="sub">{data.agenzia.sottotitolo}</div>
                      <div className="sub" style={{ marginTop: 6 }}>
                        {data.agenzia.logo ? "Logo personalizzato caricato." : "In uso il logo incorporato nell'app."}
                      </div>
                    </div>
                  </div>
                </section>

                <section className="card">
                  <header><h2>Contenuti pronti</h2><span className="sp eyebrow">da inserire quando servono</span></header>
                  <table className="tbl">
                    <tbody>
                      <tr>
                        <td>
                          <b>Procedure operative</b>
                          <div className="sub">Rilevazione settimanale, chiusura del mese, nuovo cliente. Finiscono in File come documenti dell'agenzia.</div>
                        </td>
                        <td className="r sub">{data.file.filter((f) => f.categoria === "procedura").length} presenti</td>
                        <td className="r">
                          <button className="btn ghost mini" onClick={() => upd("file", (l) => [...PROCEDURE_BASE(), ...l])}>Inserisci</button>
                        </td>
                      </tr>
                      <tr>
                        <td>
                          <b>Regole di automazione</b>
                          <div className="sub">Frequenza fuori soglia, budget in esaurimento, nuova chiusura, mese senza dati commerciali.</div>
                        </td>
                        <td className="r sub">{data.automazioni.length} presenti</td>
                        <td className="r">
                          <button className="btn ghost mini" onClick={() => upd("automazioni", (l) => [...REGOLE_BASE(), ...l])}>Inserisci</button>
                        </td>
                      </tr>
                      <tr>
                        <td>
                          <b>Aree di lavoro</b>
                          <div className="sub">Advertising, creatività, reportistica, sito e tracciamento, commerciale, amministrazione.</div>
                        </td>
                        <td className="r sub">{data.aree.length} presenti</td>
                        <td className="r">
                          <button className="btn ghost mini" onClick={() => upd("aree", (l) => [...AREE_DEFAULT.filter((a) => !l.some((x) => x.nome === a.nome)).map((a) => ({ ...a, id: uid() })), ...l])}>Completa</button>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <div className="nota">Inserire due volte crea dei doppioni: sono elementi normali, si cancellano dalla loro sezione.</div>
                </section>

                <section className="card">
                  <header><h2>Aree di lavoro</h2><span className="sp" /><button className="btn ghost mini" onClick={() => apri("aree")}>+ Area</button></header>
                  <table className="tbl">
                    <tbody>
                      {data.aree.map((a) => (
                        <tr key={a.id} className="clic" onClick={() => apri("aree", a)}>
                          <td><b>{a.nome}</b></td>
                          <td className="r sub">{data.task.filter((t) => t.areaId === a.id && t.stato !== "fatto").length} task aperti</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {!data.aree.length && <div className="empty"><b>Nessuna area</b>Le aree classificano i task: advertising, creatività, report…</div>}
                </section>
              </>
            )}

            {!attivo && vista === "task" && SezioneTask()}
            {!attivo && vista === "grafici" && SezioneGrafici()}
            {!attivo && vista === "vendite" && SezioneVendite()}
            {!attivo && vista === "chiusure" && SezioneChiusure()}
            {!attivo && vista === "automazioni" && SezioneAutomazioni()}
            {!attivo && vista === "mesi" && SezioneMesi()}
            {!attivo && vista === "semaforo" && SezioneSemaforo()}
            {!attivo && vista === "pixel" && SezionePixel()}
            {!attivo && vista === "calendario" && SezioneCalendario()}
            {!attivo && vista === "advertising" && SezioneAdv()}
            {!attivo && vista === "pianificatore" && SezionePianificatore()}
            {!attivo && vista === "funnel" && SezioneFunnel()}
            {!attivo && vista === "diario" && SezioneDiario()}
            {!attivo && vista === "economia" && SezioneEconomia()}
            {!attivo && vista === "backup" && SezioneBackup()}
            {!attivo && vista === "importa" && SezioneImporta()}
            {vista === "guida" && SezioneGuida()}
            {!attivo && vista === "academy" && SezioneAcademy()}
            {vista === "file" && SezioneFile()}
            {vista === "ai" && SezioneAI()}

            {/* ================= SOTTO-ACCOUNT ================= */}
            {attivo && vista === "cruscotto" && (
              <>
                <div className="scheda">
                  <Logo src={attivo.logo} nome={attivo.nome} size={54} />
                  <div>
                    <span className="eyebrow">{et(STATI_CLIENTE, attivo.stato)}</span>
                    <h2>{attivo.nome}</h2>
                    <div className="sub" style={{ color: "#959092", marginTop: 6 }}>
                      {attivo.fee ? `${eurDi(attivo.fee, valutaCompenso(data, attivo.id))} al mese` : "compenso non indicato"}
                      {attivo.budgetMese ? ` · budget adv ${eur(attivo.budgetMese)}/mese` : ""}
                    </div>
                  </div>
                  <div className="met">
                    <div><div className="eyebrow">Speso</div><div className="v">{eur(kpiCli.sp)}</div></div>
                    <div><div className="eyebrow">{kpiCli.rc ? "ROAS" : "CPA"}</div><div className="v">{kpiCli.rc ? `${fmt(kpiCli.roas)}×` : kpiCli.cv ? eur(kpiCli.cpa) : "—"}</div></div>
                    <div><div className="eyebrow">Avvisi</div><div className="v" style={{ color: S.avvisi.length ? "var(--acc-t)" : "#fff" }}>{S.avvisi.length}</div></div>
                  </div>
                </div>

                <div className="kpis">
                  <div className="kpi"><IcoKpi t="Conversioni" /><div className="eyebrow">Conversioni</div><div className="v">{int(kpiCli.cv)}</div><div className="s">da {S.campagne.filter((k) => ["attiva", "test"].includes(k.stato)).length} campagne attive</div></div>
                  <div className="kpi"><IcoKpi t="Lead in archivio" /><div className="eyebrow">Lead in archivio</div><div className="v">{int(S.lead.length)}</div><div className="s">{S.lead.filter((c) => ["qualificato", "opportunita"].includes(c.stato)).length} qualificati o in trattativa</div></div>
                  <div className="kpi"><IcoKpi t="Pipeline aperta" /><div className="eyebrow">Pipeline aperta</div><div className="v">{eur0(kpiCli.pipeline)}</div><div className="s">previsione ponderata {eur0(kpiCli.previsione)}</div></div>
                  <div className="kpi"><IcoKpi t="Task aperti" /><div className="eyebrow">Task aperti</div><div className="v">{S.task.filter((t) => t.stato !== "fatto").length}</div><div className="s">{S.task.filter((t) => t.stato !== "fatto" && t.scadenza && t.scadenza < oggi()).length} in ritardo</div></div>
                </div>

                {S.avvisi.length > 0 && <div className="alerts">{S.avvisi.map((a, i) => <Avviso key={i} a={a} dove={a.campagna || ""} />)}</div>}

                <div className="two">
                  <section className="card">
                    <header><h2>Dal lead al cliente</h2></header>
                    <div style={{ padding: "18px 16px 10px" }}><Funnel steps={kpiCli.funnel} /></div>
                  </section>
                  <section className="card">
                    <header><h2>Task del cliente</h2><span className="sp" /><button className="btn ghost mini" onClick={() => setVista("task")}>Apri</button></header>
                    <div className="acts">
                      {ordinaTask(S.task.filter((t) => t.stato !== "fatto")).slice(0, 6).map((t) => <RigaTask key={t.id} t={t} />)}
                      {!S.task.some((t) => t.stato !== "fatto") && <div className="empty"><b>Niente da fare</b>Nessun task aperto per questo cliente.</div>}
                    </div>
                  </section>
                </div>

                {!attivo.padreId && (
                  <section className="card">
                    <header>
                      <h2>Sottoclienti</h2>
                      <span className="sp eyebrow">
                        {haFigli ? `in ordine di cognome · ${consolidato ? "i numeri sopra li comprendono" : "esclusi dai numeri sopra"}` : "nessuno, per ora"}
                      </span>
                      <button className="btn ghost mini" onClick={() => apri("clienti", { padreId: attivo.id }, true)}>+ Sottocliente</button>
                      <button className="btn ghost mini" title="Da un foglio Google o Excel"
                        onClick={() => {
                          const padre = attivo.id;
                          setCliAtt(""); setQ(""); setConsolidato(true);
                          setImp((x) => ({ ...x, tipo: "clienti", clienteId: padre, esito: null, righe: [], intestazioni: [], mappa: {}, nomeFile: "" }));
                          setVista("importa");
                        }}>
                        Importa elenco
                      </button>
                    </header>
                    {!haFigli ? (
                      <div className="empty">
                        <b>Nessun sottocliente</b>
                        Servono quando gestisci un'azienda e insieme i suoi clienti finali: ogni sottocliente avrà campagne, funnel,
                        messaggio, ciclo di vendita, chiusure e task propri, e da qui vedrai anche i numeri consolidati del gruppo.
                      </div>
                    ) : (
                    <div className="wrap">
                      <table className="tbl">
                        <thead><tr>
                          <th>Account</th><th>Stato</th><th className="r">Campagne</th><th>Data partenza</th><th className="r">Speso</th>
                          <th className="r">Lead</th><th className="r">CPL</th><th className="r">Resa</th><th className="r">Avvisi</th><th />
                        </tr></thead>
                        <tbody>
                          {figli(attivo.id).map((z) => {
                            const p = m.portafoglio.find((x) => x.id === z.id) || { sp: 0, nCamp: 0, nAvvisi: 0, testo: "—", modo: "nd", perf: 0 };
                            /* Lead e CPL dell'ultimo mese con dati, come nella panoramica */
                            const mesiZ = righeMensili([z.id]);
                            const u = mesiZ[mesiZ.length - 1] || { lead: 0, cpl: 0 };
                            return (
                              <tr key={z.id} className="clic" onClick={() => cambiaAccount(z.id)}>
                                <td><div className="riga"><Logo src={z.logo} nome={z.nome} size={26} /><div><b>{z.nome}</b><div className="sub">{z.settore || "settore da compilare"}</div></div></div></td>
                                <td><span className="tag" data-k={z.stato}>{et(STATI_CLIENTE, z.stato)}</span></td>
                                <td className="r">{p.nCamp}</td>
                                <td>
                                  {(() => {
                                    const pa = partenza(z.id);
                                    if (!pa.data) return <span className="sub">—</span>;
                                    return (
                                      <>
                                        {dataItA(pa.data)}
                                        <div className="sub">{pa.daCampagne ? `${int(giorniTra(pa.data, oggi()))} giorni di adv` : "in gestione, senza campagne"}</div>
                                      </>
                                    );
                                  })()}
                                </td>
                                <td className="r">{p.sp ? eurDi(p.sp, z.valuta || attivo.valuta) : "—"}</td>
                                <td className="r">{u.lead ? int(u.lead) : "—"}</td>
                                <td className="r">{u.cpl ? eurDi(u.cpl, z.valuta || attivo.valuta) : "—"}</td>
                                <td className={"r " + (p.modo !== "nd" && p.perf < 1 ? "neg" : "")}>{p.testo}</td>
                                <td className="r">{p.nAvvisi ? <span className="tag" data-k="late">{p.nAvvisi}</span> : <span className="sub">ok</span>}</td>
                                <td className="r">
                                  <div className="riga" style={{ justifyContent: "flex-end", gap: 6 }}>
                                    <button className="btn ghost mini" title="Prepara la richiesta dei dati aggiornati"
                                      onClick={(e) => { e.stopPropagation(); setRichiesta({ nome: z.nome, testo: componiRichiesta(z) }); }}>Dati</button>
                                    <button className="btn ghost mini" title="Mettilo da parte senza perdere lo storico"
                                      onClick={(e) => { e.stopPropagation(); archiviaCliente(z); }}>Archivia</button>
                                    <button className="btn ghost mini" title="Cancellalo con tutti i suoi dati"
                                      onClick={(e) => { e.stopPropagation(); eliminaCliente(z); }}>Elimina</button>
                                  </div>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                    )}
                    {haFigli && (
                      <div className="nota">
                        Ogni sottocliente ha la struttura completa: campagne, funnel, messaggio, ciclo di vendita, chiusure e task suoi.
                        Con l'interruttore in alto passi dai numeri consolidati a quelli del solo account madre.
                      </div>
                    )}
                  </section>
                )}

                {(() => {
                  /* Stesse regole della sezione Advertising, in forma ridotta:
                     dalla più recente, con i filtri che servono a guardare una cosa sola. */
                  const lista = S.campagne
                    .filter((k) => (!cruStato || k.stato === cruStato)
                      && (!cruFase || faseCampagna(k).fase === cruFase)
                      && (!cruSegnale || semaforo(data, k).segnale === cruSegnale))
                    .sort((a, b) => {
                      const dA = a.inizio || "", dB = b.inizio || "";
                      if (cruOrd === "vecchie") return dA.localeCompare(dB) || a.nome.localeCompare(b.nome, "it");
                      if (cruOrd === "spesa") return b.d.sp - a.d.sp;
                      if (cruOrd === "cpl") return (a.d.cpl || Infinity) - (b.d.cpl || Infinity);
                      return dB.localeCompare(dA) || a.nome.localeCompare(b.nome, "it");
                    });
                  const attivi2 = cruStato || cruFase || cruSegnale || cruOrd !== "recenti";
                  return (
                    <section className="card">
                      <header>
                        <h2>Campagne</h2>
                        <span className="sp eyebrow">{lista.length} di {S.campagne.length}</span>
                        {attivi2 && (
                          <button className="btn ghost mini"
                            onClick={() => { setCruStato(""); setCruFase(""); setCruSegnale(""); setCruOrd("recenti"); }}>Azzera</button>
                        )}
                        <button className="btn ghost mini" onClick={() => setVista("campagne")}>Apri sezione</button>
                      </header>
                      <div className="filtri" style={{ padding: "11px 16px", borderBottom: "1px solid var(--line)" }}>
                        <select value={cruOrd} onChange={(e) => setCruOrd(e.target.value)} aria-label="Ordinamento">
                          <option value="recenti">Dalla più recente</option>
                          <option value="vecchie">Dalla più vecchia</option>
                          <option value="spesa">Spesa, dalla più alta</option>
                          <option value="cpl">CPL, dal più basso</option>
                        </select>
                        <select value={cruStato} onChange={(e) => setCruStato(e.target.value)} aria-label="Filtra per stato">
                          <option value="">Tutti gli stati</option>
                          {STATI_CAMPAGNA.map((x) => <option key={x.v} value={x.v}>{x.l}</option>)}
                        </select>
                        <select value={cruFase} onChange={(e) => setCruFase(e.target.value)} aria-label="Filtra per fase">
                          <option value="">Tutte le fasi</option>
                          <option value="test">Test</option>
                          <option value="matura">Matura</option>
                          <option value="attiva">Attiva</option>
                          <option value="chiusa">Chiuse</option>
                          <option value="errore">Ferme (pagamento o account)</option>
                        </select>
                        <select value={cruSegnale} onChange={(e) => setCruSegnale(e.target.value)} aria-label="Filtra per semaforo">
                          <option value="">Tutti i segnali</option>
                          {SEGNALI.map((g) => <option key={g.v} value={g.v}>{g.l}</option>)}
                        </select>
                      </div>
                      <div className="wrap"><table className="tbl"><TestaCampagne /><tbody>{lista.map((k) => <RigaCampagna key={k.id} k={k} />)}</tbody></table></div>
                      {!S.campagne.length && <div className="empty"><b>Nessuna campagna</b>Aggiungine una dalla sezione Advertising.</div>}
                      {!!S.campagne.length && !lista.length && <div className="empty"><b>Nessuna campagna con questi filtri</b>Premi Azzera per rivederle tutte.</div>}
                    </section>
                  );
                })()}
              </>
            )}

            {attivo && vista === "contatti" && (
              <section className="card">
                <div className="wrap">
                  <table className="tbl">
                    <thead><tr><th>Contatto</th><th>Tipo</th><th>Stato</th><th>Provenienza</th><th>Recapiti</th></tr></thead>
                    <tbody>
                      {S.contatti.filter((c) => cerca(c.nome + " " + c.azienda + " " + c.email)).map((c) => (
                        <tr key={c.id} className="clic" onClick={() => apri("contatti", c)}>
                          <td><b>{c.nome}</b><div className="sub">{[c.azienda, c.ruolo].filter(Boolean).join(" · ")}</div></td>
                          <td><span className="tag" data-k={c.tipo === "referente" ? "attivo" : ""}>{et(TIPI_CONTATTO, c.tipo)}</span></td>
                          <td><span className="tag" data-k={c.stato}>{et(STATI_LEAD, c.stato)}</span></td>
                          <td className="sub">{c.campagnaId ? (S.campagne.find((k) => k.id === c.campagnaId) || {}).nome || "—" : "Diretto"}</td>
                          <td className="sub">{c.email}<br />{c.telefono}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {!S.contatti.length && <div className="empty"><b>Nessun contatto</b>Inserisci i lead di questo cliente e i suoi referenti interni.</div>}
              </section>
            )}

            {attivo && vista === "pipeline" && (
              <>
                <div className="board">
                  {FASI_APERTE.map((f) => {
                    const items = S.trattative.filter((t) => t.fase === f.v);
                    return (
                      <div className="col" key={f.v}>
                        <h3>{f.l}<span>{eur0(items.reduce((s, t) => s + t.valore, 0))}</span></h3>
                        <div className="list">
                          {items.map((t) => {
                            const i = FASI.findIndex((x) => x.v === t.fase);
                            const c = contatto(t.contattoId);
                            return (
                              <article className="deal" key={t.id}>
                                <b>{t.titolo}</b>
                                <div className="sub">{c ? c.azienda || c.nome : "senza contatto"} · {dataIt(t.chiusura)}</div>
                                <div className="row">
                                  <span className="val">{eur(t.valore)}</span>
                                  <span style={{ marginLeft: "auto", display: "flex", gap: 4 }}>
                                    <button className="move" onClick={() => muovi(t, -1)} disabled={i === 0} aria-label="Fase precedente">‹</button>
                                    <button className="move" onClick={() => muovi(t, 1)} aria-label="Fase successiva">›</button>
                                    <button className="move" onClick={() => apri("trattative", t)} aria-label="Modifica">✎</button>
                                  </span>
                                </div>
                              </article>
                            );
                          })}
                          {!items.length && <div className="sub" style={{ padding: "8px 2px" }}>Vuota.</div>}
                        </div>
                      </div>
                    );
                  })}
                </div>
                <section className="card">
                  <header><h2>Chiuse</h2></header>
                  <table className="tbl">
                    <thead><tr><th>Trattativa</th><th>Esito</th><th>Origine</th><th className="r">Valore</th></tr></thead>
                    <tbody>
                      {S.trattative.filter((t) => ["vinto", "perso"].includes(t.fase)).map((t) => (
                        <tr key={t.id} className="clic" onClick={() => apri("trattative", t)}>
                          <td><b>{t.titolo}</b></td>
                          <td><span className="tag" data-k={t.fase}>{faseDi(t.fase).l}</span></td>
                          <td className="sub">{(S.campagne.find((k) => k.id === t.campagnaId) || {}).nome || "—"}</td>
                          <td className="r">{eur(t.valore)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {!S.trattative.some((t) => ["vinto", "perso"].includes(t.fase)) && <div className="empty"><b>Ancora nulla di chiuso</b>Le trattative concluse finiscono qui.</div>}
                </section>
              </>
            )}

            {attivo && vista === "messaggio" && SezioneMessaggio()}
            {attivo && vista === "mesi" && SezioneMesi()}
            {attivo && vista === "semaforo" && SezioneSemaforo()}
            {attivo && vista === "pixel" && SezionePixel()}
            {attivo && vista === "calendario" && SezioneCalendario()}
            {attivo && vista === "campagne" && SezioneAdv()}
            {attivo && vista === "winner" && SezioneWinner()}

            {attivo && vista === "avvisi" && (
              <div className="alerts">
                <button className="btn ghost mini" style={{ alignSelf: "flex-start" }} onClick={() => setVista("cruscotto")}>‹ Torna al cruscotto</button>
                {[...S.avvisi].sort((a, b) => (a.gravita === b.gravita ? 0 : a.gravita === "critica" ? -1 : 1)).map((a, i) => <Avviso key={i} a={a} dove={a.campagna || ""} />)}
                {!S.avvisi.length && <section className="card"><div className="empty"><b>Nessuna soglia superata</b>Le campagne di {attivo.nome} sono dentro i benchmark.</div></section>}
              </div>
            )}

            {attivo && vista === "task" && SezioneTask()}
            {attivo && vista === "grafici" && SezioneGrafici()}
            {attivo && vista === "vendite" && SezioneVendite()}
            {attivo && vista === "chiusure" && SezioneChiusure()}
            {attivo && vista === "pianificatore" && SezionePianificatore()}
            {attivo && vista === "funnel" && SezioneFunnel()}
            {attivo && vista === "diario" && SezioneDiario()}
            {attivo && vista === "report" && SezioneReport()}

            {attivo && vista === "scheda" && (
              <>
                {PannelloMetaCliente(attivo)}
                <section className="card">
                  <header><h2>Dati dell'account</h2></header>
                  <div style={{ padding: 18, display: "flex", alignItems: "center", gap: 16, borderBottom: "1px solid var(--line)" }}>
                    <Logo src={attivo.logo} nome={attivo.nome} size={64} />
                    <div>
                      <div style={{ fontSize: 18, fontWeight: 650 }}>{attivo.nome}</div>
                      <div className="sub">{attivo.logo ? "Logo caricato" : "Nessun logo: mostra le iniziali"}</div>
                    </div>
                  </div>
                  <table className="tbl"><tbody>
                    <tr><td className="sub">Settore</td><td>{attivo.settore || "—"}</td></tr>
                    <tr><td className="sub">Sede</td><td>{attivo.citta || "—"}</td></tr>
                    <tr><td className="sub">Stato</td><td><span className="tag" data-k={attivo.stato}>{et(STATI_CLIENTE, attivo.stato)}</span></td></tr>
                    <tr><td className="sub">Compenso mensile</td><td className="mono">{attivo.fee ? eurDi(attivo.fee, valutaCompenso(data, attivo.id)) : "—"}</td></tr>
                    <tr><td className="sub">Budget adv mensile</td><td className="mono">{attivo.budgetMese ? eur(attivo.budgetMese) : "—"}</td></tr>
                    <tr>
                      <td className="sub">Margine e pareggio</td>
                      <td className="mono">
                        {attivo.margineLordo
                          ? `${fmt(attivo.margineLordo, 0)}% lordo${attivo.resi ? `, ${fmt(attivo.resi, 0)}% resi` : ""} · ROAS di pareggio ${fmt(roasPareggio(attivo))}×`
                          : <span className="neg">non impostato: il ROAS resta un numero senza significato economico</span>}
                      </td>
                    </tr>
                    <tr>
                      <td className="sub">In gestione dal</td>
                      <td>
                        {attivo.creato ? dataItA(attivo.creato) : <span className="sub">da impostare</span>}
                        {attivo.creato && <span className="sub" style={{ marginLeft: 8 }}>{int(giorniTra(attivo.creato, oggi()))} giorni</span>}
                      </td>
                    </tr>
                    <tr>
                      <td className="sub">Valuta</td>
                      <td>
                        {valutaCliente(data, attivo.id)}
                        {!attivo.valuta && <span className="sub" style={{ marginLeft: 8 }}>{attivo.padreId && valutaCliente(data, attivo.padreId) !== (data.agenzia.valuta || "EUR") ? "dal cliente madre" : "dall'agenzia"}</span>}
                        <span className="sub" style={{ marginLeft: 8 }}>{eurDi(1234.56, valutaCliente(data, attivo.id))}</span>
                      </td>
                    </tr>
                    <tr>
                      <td className="sub">Data contratto</td>
                      <td>
                        {attivo.dataContratto ? dataItA(attivo.dataContratto) : <span className="sub">non indicata</span>}
                        {attivo.dataContratto && attivo.creato && attivo.dataContratto !== attivo.creato && (
                          <span className="sub" style={{ marginLeft: 8 }}>
                            {attivo.dataContratto < attivo.creato
                              ? `firmato ${int(giorniTra(attivo.dataContratto, attivo.creato))} giorni prima dell'avvio`
                              : `firmato ${int(giorniTra(attivo.creato, attivo.dataContratto))} giorni dopo l'avvio`}
                          </span>
                        )}
                      </td>
                    </tr>
                  </tbody></table>
                </section>
                {attivo.note && <section className="card"><header><h2>Note e vincoli</h2></header><div style={{ padding: "14px 16px" }}>{attivo.note}</div></section>}

                <section className="card">
                  <header>
                    <h2>Dati di fatturazione</h2>
                    <span className="sp" />
                    <button className="btn ghost mini" onClick={() => {
                      const blocco = [
                        attivo.ragioneSociale || attivo.nome,
                        [attivo.indirizzo, [attivo.cap, attivo.cittaFatt, attivo.provincia ? `(${attivo.provincia})` : ""].filter(Boolean).join(" ")].filter(Boolean).join("\n"),
                        attivo.paese && attivo.paese !== "Italia" ? attivo.paese : "",
                        attivo.piva ? `P. IVA ${attivo.piva}` : "",
                        attivo.codiceFiscale && attivo.codiceFiscale !== attivo.piva ? `C.F. ${attivo.codiceFiscale}` : "",
                        attivo.sdi ? `Codice destinatario ${attivo.sdi}` : "",
                        attivo.pec ? `PEC ${attivo.pec}` : "",
                        attivo.termini ? `Pagamento: ${et(TERMINI_PAGAMENTO, attivo.termini)}${attivo.metodo ? ` · ${et(METODI_PAGAMENTO, attivo.metodo)}` : ""}` : "",
                      ].filter(Boolean).join("\n");
                      chiedi({
                        titolo: "Dati per la fattura",
                        testo: blocco,
                        etichetta: "Chiudi",
                        azione: () => { try { navigator.clipboard.writeText(blocco); } catch (e) { /* selezione manuale */ } },
                      });
                    }}>Vedi e copia</button>
                    <button className="btn ghost mini" onClick={() => apri("clienti", attivo)}>Modifica</button>
                  </header>
                  <div className="wrap">
                    <table className="tbl"><tbody>
                      <tr><td className="sub">Ragione sociale</td><td>{attivo.ragioneSociale || <span className="neg">da compilare</span>}</td></tr>
                      <tr>
                        <td className="sub">Partita IVA</td>
                        <td className="mono">
                          {attivo.piva || <span className="neg">da compilare</span>}
                          {attivo.piva && !/^\d{11}$/.test(String(attivo.piva).replace(/\s|IT/gi, "")) &&
                            <div className="sub neg">non ha 11 cifre: controlla</div>}
                        </td>
                      </tr>
                      <tr><td className="sub">Codice fiscale</td><td className="mono">{attivo.codiceFiscale || "—"}</td></tr>
                      <tr>
                        <td className="sub">Sede legale</td>
                        <td>{[attivo.indirizzo, [attivo.cap, attivo.cittaFatt, attivo.provincia && `(${attivo.provincia})`].filter(Boolean).join(" "), attivo.paese].filter(Boolean).join(" · ") || "—"}</td>
                      </tr>
                      <tr>
                        <td className="sub">Fattura elettronica</td>
                        <td className="mono">
                          {attivo.sdi || attivo.pec
                            ? [attivo.sdi && `SDI ${attivo.sdi}`, attivo.pec].filter(Boolean).join(" · ")
                            : <span className="neg">manca SDI o PEC: la fattura non arriva</span>}
                        </td>
                      </tr>
                      <tr><td className="sub">Amministrazione</td><td>{[attivo.referenteAmm, attivo.emailAmm].filter(Boolean).join(" · ") || "—"}</td></tr>
                      <tr>
                        <td className="sub">Pagamento</td>
                        <td>{et(TERMINI_PAGAMENTO, attivo.termini)} · {et(METODI_PAGAMENTO, attivo.metodo)}
                          {attivo.fee > 0 && <div className="sub">{eurDi(attivo.fee, valutaCompenso(data, attivo.id))} al mese ricorrenti</div>}</td>
                      </tr>
                    </tbody></table>
                  </div>
                </section>
                <section className="card">
                  <header><h2>Cosa contiene questo account</h2></header>
                  <table className="tbl"><tbody>
                    <tr><td>Contatti</td><td className="r">{S.contatti.length}</td></tr>
                    <tr><td>Trattative</td><td className="r">{S.trattative.length}</td></tr>
                    <tr><td>Campagne</td><td className="r">{S.campagne.length}</td></tr>
                    <tr><td>Task</td><td className="r">{S.task.length}</td></tr>
                  </tbody></table>
                </section>

                <section className="card">
                  <header>
                    <h2>Installazione del metodo</h2>
                    <span className="sp eyebrow">
                      {PASSI_AOA.filter((p) => (attivo.aoa || {})[p.k]).length} di {PASSI_AOA.length} completati
                    </span>
                  </header>
                  <div className="barra" style={{ margin: "0 16px", height: 6 }}>
                    <i style={{ width: `${(PASSI_AOA.filter((p) => (attivo.aoa || {})[p.k]).length / PASSI_AOA.length) * 100}%` }} />
                  </div>
                  <div className="acts">
                    {PASSI_AOA.map((p, i) => {
                      const fatto = !!(attivo.aoa || {})[p.k];
                      return (
                        <div className="tkrow" key={p.k} data-done={fatto ? "1" : "0"}>
                          <input type="checkbox" checked={fatto} aria-label={p.l}
                            onChange={(e) => upd("clienti", (l) => l.map((z) => (z.id === attivo.id
                              ? { ...z, aoa: { ...(z.aoa || {}), [p.k]: e.target.checked } } : z)))} />
                          <div className="t">
                            <b>{i + 1}. {p.l}</b>
                            <div className="sub" style={{ fontSize: 12.6, marginTop: 3, maxWidth: "78ch" }}>{p.d}</div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                  <div className="nota">
                    I primi cinque passi sono lavoro da fare prima di pubblicare: nessun annuncio ha mai guadagnato un euro prima della pubblicazione,
                    ma è lì che si decide se guadagnerà. Gli ultimi due sono i sistemi che tengono viva la campagna dopo.
                  </div>
                </section>

                <section className="card">
                  <header>
                    <h2>Archiviare il cliente</h2>
                    <span className="sp eyebrow">reversibile, senza perdere nulla</span>
                    <button className="btn ghost mini" onClick={() => archiviaCliente(attivo)}>Archivia</button>
                  </header>
                  <div className="nota" style={{ paddingTop: 14 }}>
                    Usalo quando smetti di seguire un cliente ma potrebbe tornare: esce dall'elenco, dai totali e dagli avvisi,
                    e tutto il suo storico finisce in un deposito da cui puoi riprenderlo quando vuoi.
                  </div>
                </section>

                {data.deposito.some((v2) => v2.clienteId === attivo.id && v2.motivo !== "advertising") && (
                  <section className="card">
                    <header>
                      <h2>Storico in deposito</h2>
                      <span className="sp eyebrow">
                        {data.deposito.filter((v2) => v2.clienteId === attivo.id && v2.motivo !== "advertising")
                          .reduce((s2, v2) => s2 + COLL_CLI.reduce((x, c) => x + (v2.dati[c] || []).length, 0), 0)} voci
                      </span>
                      <button className="btn" onClick={() => chiedi({
                        titolo: "Reinserire lo storico?",
                        testo: "Riporto dentro campagne, fasi, funnel, periodi commerciali, chiusure, task e file archiviati in precedenza. I dati inseriti dopo la riattivazione restano dove sono.",
                        etichetta: "Reinserisci",
                        azione: () => riportaStorico(attivo.id),
                      })}>Reinserisci lo storico</button>
                    </header>
                    <div className="nota" style={{ paddingTop: 14 }}>
                      Quando hai riattivato questo cliente hai scelto di ripartire da zero. I dati precedenti sono conservati qui
                      e non entrano in nessun calcolo finché non li reinserisci.
                    </div>
                  </section>
                )}

                <section className="card">
                  <header><h2>Eliminare il cliente</h2><span className="sp eyebrow">operazione irreversibile</span></header>
                  <div style={{ padding: "16px 18px", fontSize: 13.3, color: "var(--ink2)", maxWidth: "72ch" }}>
                    Elimina {attivo.nome} e tutto ciò che gli appartiene: contatti, trattative, campagne, ad set, periodi commerciali,
                    chiusure, funnel, task, annotazioni, esperimenti, file e storico delle rilevazioni. Prima di procedere ti mostro
                    l'elenco esatto di quanto stai per perdere.
                    <br /><br />
                    Se il rapporto è finito ma i dati ti servono ancora — per uno storico, per una causa, per riprendere il cliente fra un anno —
                    non eliminarlo: mettilo su "Chiuso" dalla scheda. Sparisce dal semaforo e dai conteggi attivi, ma resta consultabile.
                  </div>
                  <div style={{ padding: "0 18px 18px", display: "flex", gap: 8, flexWrap: "wrap" }}>
                    <button className="btn ghost" onClick={esportaBackup}>Scarica un backup adesso</button>
                    <button className="btn danger" onClick={() => eliminaCliente(attivo)}>Elimina {attivo.nome}</button>
                  </div>
                </section>
              </>
            )}
          </div>
        </div>
      </div>

      {ed && (
        <Editor
          titolo={ed.tipo === "agenzia" ? "Identità dell'agenzia"
            : (ed.nuovo ? "Nuovo · " : "Modifica · ") + ({ clienti: "cliente", contatti: "contatto", trattative: "trattativa", campagne: "campagna", task: "task", aree: "area", vendite: "periodo commerciale", chiusure: "chiusura", automazioni: "automazione", creativi: "ad set o annuncio", annotazioni: "annotazione", esperimenti: "esperimento", funnel: "funnel", messaggio: "piano di messaggio", periodi: "fase della campagna", modelli: "modello di percorso", eventi: "evento", categorie: "categoria", file: "documento", giornate: "giornata", rilevazioni: "rilevazione" }[ed.tipo] || ed.tipo)}
          sotto={["contatti", "trattative", "campagne"].includes(ed.tipo) && attivo ? attivo.nome : ""}
          fields={CAMPI[ed.tipo]}
          draft={ed.draft}
          setDraft={(d) => {
            if (ed.tipo === "campagne" && d.obiettivo !== ed.draft.obiettivo) {
              const b = BMK_DEFAULT[d.obiettivo];
              d = { ...d, roasMin: perInput(b.roasMin), cpaMax: perInput(b.cpaMax), cplMax: perInput(b.cplMax), ctrMin: perInput(b.ctrMin), freqMax: perInput(b.freqMax), cpmMax: perInput(b.cpmMax) };
            }
            if (ed.tipo === "funnel" && d.modello !== ed.draft.modello && d.modello) {
              const mo = optModelli.find((x) => x.v === d.modello);
              if (mo && mo.passi.length) {
                d = { ...d, passi: mo.passi.map(([nome, tipo]) => ({ id: uid(), nome, tipo, visite: "", conversioni: "" })) };
              }
            }
            setEd({ ...ed, draft: d });
          }}
          onSave={salva}
          liste={{
            linee: [...new Set(data.campagne.map((k) => k.linea).filter(Boolean))].sort(),
            campagne: data.campagne.filter((k) => k.archiviata !== "si").map((k) => ({
              id: k.id, nome: k.nome, clienteId: k.clienteId, stato: etStato(k),
              altro: k.funnelId && ed && k.funnelId !== ed.draft.id
                ? (data.funnel.find((f) => f.id === k.funnelId) || {}).nome : "",
            })).sort((a, b) => a.nome.localeCompare(b.nome, "it")),
          }}
          onClose={() => { setNuovoPadre(""); setEd(null); }}
          onDelete={ed.nuovo || ed.tipo === "agenzia" ? null : elimina}
        />
      )}

      {richiesta && (
        <div className="scrim" onMouseDown={(e) => e.target === e.currentTarget && setRichiesta(null)}>
          <div className="modal" role="dialog" aria-modal="true" aria-label="Richiesta di aggiornamento" style={{ maxWidth: 620 }}>
            <header><h2>Richiesta pronta · {richiesta.nome}</h2></header>
            <div style={{ padding: "16px 18px", display: "grid", gap: 10 }}>
              <span className="hint">
                Copia questo testo e incollalo nella chat con il consulente: contiene account, periodo mancante ed elenco
                delle campagne, così i file tornano già pronti da importare.
              </span>
              <textarea readOnly rows={14} value={richiesta.testo}
                style={{ fontFamily: "var(--mono)", fontSize: 12, lineHeight: 1.5 }}
                onFocus={(e) => e.target.select()} />
            </div>
            <footer>
              <span className="sp" />
              <button className="btn ghost" onClick={() => setRichiesta(null)}>Chiudi</button>
              <button className="btn" onClick={() => {
                const t = document.querySelector('.modal textarea');
                if (t) { t.select(); try { document.execCommand("copy"); } catch (e) { /* selezionato comunque */ } }
                if (navigator.clipboard) navigator.clipboard.writeText(richiesta.testo).catch(() => {});
              }}>Copia</button>
            </footer>
          </div>
        </div>
      )}
      {conferma && <Conferma c={conferma} onChiudi={() => setConferma(null)} />}
    </div>
  );
}

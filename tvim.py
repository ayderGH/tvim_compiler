import argparse
import yaml
import shutil
import os
import subprocess
import re
import logging


logger = logging.getLogger('tvim')
logger.setLevel(logging.INFO)


def get_text_between_braces(text, open_pos=0):
    """
    Получить текст между фигурными скобками.
    """
    counter = 1
    i = open_pos
    while counter > 0 and i < len(text):
        if text[i] == '{':
            counter += 1
        elif text[i] == '}':
            counter -= 1
        if counter == 0:
            return text[open_pos:i], i
        i += 1


class Article:
    """
    Объектная модель статьи журнала.
    """
    def __init__(self, path):
        self.path = path
        self.title = {}
        self.authors = {}
        self.author_details = []
        self.udc = None
        self.msc2010 = None
        self.abstracts = {}
        self.sections = {}
        self.bibliography = {}
        self.text = self.get_text()
        self.article_text = None

    def get_text(self):
        tex_file = [f for f in os.listdir(self.path)
                    if f.endswith('.tex')]
        if tex_file:
            tex_file = os.path.join(self.path, tex_file[0])
            with open(tex_file, 'rt') as f:
                text = f.read()
        return text

    @staticmethod
    def normalize_text(text):
        """
        Удалить:
            textbf
            textit
        """
        text = re.sub(r'\\textbf|\\textit|\\it|\\bf', '', text)
        text = re.sub(r'[\{|\}]', '', text)
        text = re.sub(r'\s{2,}', ' ', text)
        return text.strip()

    def extract_title(self):
        """
        Извлечь заголовок.
        """
        text = self.text
        self.title = {}
        title_pattern = r'\\title\{(?P<title>.*)\}'
        m = re.search(title_pattern, text)
        if m:
            title = m['title']
            # удалить, если имеется, 'footnote'
            title = re.sub(r'\\footnote\{.*\}', '', title)
            self.title['ru'] = title
        else:
            logger.error('Не найден заголовок в {}!'.format(self.path))

    def extract_authors(self):
        """
        Извлечь авторов.
        """
        text = self.text
        self.authors = []
        for m1 in re.finditer(r'\\author\{(.*)\}', text):
            authors = m1[1]
            authors = re.sub(r'\\[;,.:]+', ' ', authors)
            authors = re.sub(r'\s{2,}', ' ', authors)
            name_abbr_pattern = r'(?P<name>[A-ZА-ЯЁa-zа-яё]{1,2}\.)'
            patronymic_abbr_pattern = r'(?P<patronymic>[A-ZА-ЯЁa-zа-яё]{1,2}\.)'
            family_pattern = r'(?P<family>[A-ZА-Яa-zа-яё]{2,})'
            patterns = [r'{}\s*{}\s*{}'.format(name_abbr_pattern,
                                               patronymic_abbr_pattern,
                                               family_pattern),
                        r'{}\s*{}\s*{}'.format(family_pattern,
                                               name_abbr_pattern,
                                               patronymic_abbr_pattern)]

            for pattern in patterns:
                for m in re.finditer(pattern, authors):
                    self.authors.append({'family': m['family'],
                                         'name': m['name'],
                                         'patronymic': m['patronymic']})
        if not self.authors:
            logger.error('Не найдены авторы в {}!'.format(self.path))

    def extract_ru_abstracts(self):
        """
        Извлечь русскую аннотацию.
        """
        text = self.text
        self.abstracts['ru'] = ''
        begin_pattern = r'\\begin\{abstractXr\}.*'
        m = re.search(begin_pattern, text)
        if m:
            p0 = m.end()
            end_pattern = r'\\end\{abstractXr\}'
            m = re.search(end_pattern, text)
            if m:
                p1 = m.start()
                abstract = text[p0:p1]
                self.abstracts['ru'] = abstract
        if not self.abstracts['ru']:
            logger.error('Не найдена русская аннотация в {}!'.format(self.path))

    def extract_en_abstracts(self):
        """
        Извлечь английскую аннотацию.
        """
        self.abstracts['en'] = ''
        text = self.text
        begin_pattern = r'\\begin\{abstractX\}.*'
        m = re.search(begin_pattern, text)
        if m:
            p0 = m.end()
            end_pattern = r'\\end\{abstractX\}'
            m = re.search(end_pattern, text)
            if m:
                p1 = m.start()
                abstract = text[p0:p1]
                self.abstracts['en'] = abstract
        if not self.abstracts['en']:
            logger.error('Не найдена английская аннотация в {}!'.
                         format(self.path))

    def extract_sections(self):
        """
        Извлечь разделы (главы) статьи.
        """
        text = self.text
        pattern = r'\\section\*?\{'
        self.sections = []
        for m in re.finditer(pattern, text):
            t = get_text_between_braces(text, m.end())
            if t:
                self.sections.append(self.normalize_text(t[0]))
        if not self.sections:
            logger.error('Не найдены разделы в {}!'.format(self.path))
        else:
            print(self.sections)

    def extract_bibliography(self):
        """
        Извлечь список литературы.
        """
        # russian
        # text = self.text
        # begin_pattern = r'\\begin\{thebibliography\}.*'
        # m = re.search(begin_pattern, text)
        # if m:
        #     p0 = m.end()
        #     end_pattern = r'\\end\{thebibliography\}'
        #     m = re.search(end_pattern, text)
        #     if m:
        #         p1 = m.start()
        #         abstract = text[p0:p1]
        #         return re.sub(r'\%.*', '', abstract)

    def extract_udc(self):
        """
        Извлечь УДК.
        """
        p = r'(?<=УДК\:)\s*(.*)(?=\})'
        m = re.search(p, self.text)
        if m:
            self.udc = m[1]
        else:
            self.udc = '???'
            logger.error('Не найден УДК в {}!'.format(self.path))

    def extract_msc2010(self):
        """
        Извлечь MSC2010.
        """
        p = r'(?<=MSC2010\:)\s*(.*)(?=\})'
        m = re.search(p, self.text)
        if m:
            self.msc2010 = m[1]
        else:
            self.msc2010 = '???'
            logger.error('Не найден MSC2010 в {}!'.format(self.path))

    def update_image_path(self):
        """
        Обновить пути к файлам изображений.
        """
        text = self.article_text
        pos = []
        for m in re.finditer(r'\\includegraphics.*?\{(.+?)\}', text):
            graph_text = text[m.start():m.end()]
            m_image_name = re.search(r'{(.*)}', graph_text)
            pos.append((m.start() + m_image_name.start(),
                        m.start() + m_image_name.end()))
        d = len(self.path) + 1
        for i, p in enumerate(pos):
            image_name = text[p[0] + i*d+1:p[1] + i*d-1]
            text = text[:p[0] + i*d] + \
                   '{{{}/{}}}'.format(self.path, image_name) + \
                   text[p[1] + i*d:]
        self.article_text = text

    def add_content_lines(self):
        ru_con = '\\addcontentsline{{toc}}{{art}}' \
                 '{{\\textbf{{{authors}}} {title}}}'.\
                 format(authors=', '.join(['{} {} {}'.format(a['family'],
                                                     a['name'],
                                                     a['patronymic'])
                                          for a in self.authors]),
                        title=self.title['ru'])

        en_con = ''
        # en_con = '\\addcontentsline{{toc}}{{art}}' \
        #          '{{\\textbf{{{authors}}}{title}}}'.\
        #          format(authors=self.authors['en'],
        #                 title=self.title['en'])
        return '{}\n{}\n'.format(ru_con, en_con)

    def extract_author_details(self):
        p = r'\\authorInfo\{.*\}?'
        self.author_details = []
        for m in re.finditer(p, self.text):
            self.author_details.append(self.text[m.start():m.end()])

    def parse(self):
        self.extract_title()
        self.extract_authors()
        self.extract_author_details()
        self.extract_ru_abstracts()
        self.extract_en_abstracts()
        self.extract_udc()
        self.extract_msc2010()
        self.extract_sections()
        self.extract_bibliography()

    def compile(self):
        """
        Скомпилировать статью.
            1. Выделить необходимую информацию: заголовки, список авторов,
               аннотации и т.д.
            2. Выделить "чистый" текст.
            2. Добавить служебную информацию.
        """
        self.parse()
        m_start = re.search(r'\\markboth', self.text)
        m_end = re.search(r'\\end{thebibliography}', self.text)
        if m_start and m_end:
            self.article_text = self.text[m_start.start():m_end.end()]
            self.article_text = '\input{__init_counters__}\n' + \
                                '\input{__to_rus__}\n\n' + \
                                self.add_content_lines() + \
                                self.article_text
            article_path = os.path.join(self.path, '__article.tex')
            self.update_image_path()
            with open(article_path, 'wt') as f:
                f.write(self.article_text)
            return article_path
        else:
            raise Exception('Incorrect article {}'.format(self.path))


class TvimDocument:
    """
    Объектная модель выпуска журнала.
    """
    def __init__(self, config):
        self.config = config
        self.articles = []
        # parameters
        self.year = self.config['tvim']['year']
        self.number = self.config['tvim']['number']
        self.total_number = self.config['tvim']['total number']
        self.protocol_number = self.config['tvim'].get('protocol number', '???')
        self.protocol_day = self.config['tvim'].get('protocol day', '???')
        self.protocol_month = self.config['tvim'].get('protocol month', '???')
        self.protocol_monthname = self.config['tvim'].get('protocol month name',
                                                          '???')
        self.protocol_year = self.config['tvim'].get('protocol year', '???')
        self.resources = self.config['path']['resources']
        self.page_count = 0

        self.root_path = 'numbers/tvim_{}_{}'.format(self.year, self.number)

    @classmethod
    def from_config(cls, path):
        """
        Parameters
        ----------
            path: str
                Путь к конфигурационному файлу
        """
        with open(path, 'rt') as config_file:
            config = yaml.load(config_file, Loader=yaml.SafeLoader)
        return cls(config)

    def _update_params(self):
        """
        Обновляет основные параметры выпуска журнала, такие как год, номер,
        протокол, даты выхода в свет и др.
        """
        # обновление параметров на русском языке
        params = [
            '\def\\tvimname{Таврический вестник информатики и математики}\n',
            '\def\\tvimnumber{{№\,{number}\,({total_number})}}\n'.format(
                number=self.number, total_number=self.total_number),
            '\def\\tvimyear{{{year}}}\n'.format(year=self.year),
            '\def\\tvimemail{article@tvim{.}info}\n',
            '\def\\tvimwww{www{.}tvim{.}info}\n',
            '\def\protocolnumber{{{}}}'.format(self.protocol_number),
            '\def\protocolday{{{}}}\n'.format(self.protocol_day),
            '\def\protocolmonthname{{{}}}\n'.format(self.protocol_monthname),
            '\def\protocolmonth{{{}}}\n'.format(self.protocol_month),
            '\def\protocolyear{{{}}}\n'.format(self.protocol_year),
            '\def\protocol{№\,\protocolnumber\ от~\protocolday~'
            '\protocolmonthname~\protocolyear\,г.}\n',
            '\def\sign2print{\protocolday.\protocolmonth.\protocolyear}\n',
            '\def\print_page_count{{{}}}\n'.format(round(self.page_count
                                                         * 0.1056), 1),
            '\def\\tvimissn{ISSN\;1729-3901}\n',
            '\\newlength{\myparindent}\n',
            '\\newlength{\myinter}\n'
            ]

        with open('__params__.tex', 'wt') as f:
            f.writelines(params)

        # обновление параметров на английском языке
        params = [
            '\def\\tvimnameen{Taurida Journal of~Computer Science Theory and~Mathematics}\n',
            '\def\\tvimnumberen{{{}}}\n'.format(self.number),
            '\def\\tvimnumberwithtotalen{{No.\;{}\;({})}}\n'.format(
                self.number, self.total_number),
            '\def\\tvimyearen{{{}}}\n'.format(self.year),
            '\def\\tvimemailen{article@tvim{.}info}\n',
            '\def\\tvimwwwen{www{.}tvim{.}info}\n',
            '\def\protocolen{{No.\,? from {}/{}/{}.}}\n'.format(
                self.protocol_month, self.protocol_day, self.protocol_year),
            '\def\Profen{Professor}\n',
            '\def\Docenten{Associate professor}\n',
            '\def\Dfmnen{Doctor of Physico-Mathematical Sciences}\n',
            '\def\Dtnen{Doctor of Engineering Sciences}\n',
            '\def\Kfmnen{Candidate of Physico-Mathematical Sciences}\n',
            '\def\profen{professor}\n',
            '\def\docenten{associate professor}\n',
            '\def\dfmnen{doctor of Physico-Mathematical Sciences}\n',
            '\def\dtnen{doctor of Engineering Sciences}\n',
            '\def\kfmnen{candidate of Physico-Mathematical Sciences}\n',
            ]

        with open('__params_en__.tex', 'wt') as f:
            f.writelines(params)

    def _build(self):
        articles_path = os.path.join('articles')
        articles = [f for f in os.listdir(articles_path)
                    if not f.startswith('.')]
        art_file_content = []
        author_details = []
        referats = []
        for art in articles:
            article = Article(path=os.path.join(articles_path, art))
            self.articles.append(article)
            art_path = article.compile()
            author_details.extend(article.author_details)
            if art_path:
                art_file_content.append(art_path)
            with open('articles.tex', 'wt') as f:
                f.writelines(['\input{{{}}}\n'.format(art_path)
                              for art_path in art_file_content])

        author_details = sorted(author_details)
        with open('authors.tex', 'at') as authors_file:
            authors_file.writelines('\n\n\medskip\n'.join(author_details))


    def update_article(self, article):
        pass

    def compile(self):
        print('Компиляция ТВИМ {year} #{number}. '
              'Пожалуйста подождите'.format(year=self.year,
                                            number=self.number))
        # задаем корневую папку и копируем необходимые файлы
        # для компиляции журнала
        if os.path.exists(self.root_path):
            shutil.rmtree(self.root_path)
        shutil.copytree(self.resources, self.root_path)
        shutil.copytree(self.config['path']['articles'],
                        os.path.join(self.root_path, 'articles'))

        cur_dir = os.path.curdir
        try:
            os.chdir(self.root_path)
            main_filename = 'tvim_{}_{}.tex'.format(self.year, self.number)
            os.rename('tvim_main.tex', main_filename)
            self._update_params()
            self._build()
            cmd = ['pdflatex', 'tvim_{}_{}.tex'.format(self.year, self.number)]
            res = subprocess.run(cmd)
            if res.returncode == 0:
                # второй запуск для создания ссылок и содержания
                subprocess.run(cmd)
                print("ok")
            else:
                print("ERROR: something is wrong. Please look at log files.")
        finally:
            os.chdir(cur_dir)


if __name__ == '__main__':
    argparser = argparse.ArgumentParser(description="TVIM compiler")
    argparser.add_argument("--config", "-C", type=str,
                           default="configs/config.yaml",
                           help="path to config file in YAML format")
    args = argparser.parse_args()

    tvim = TvimDocument.from_config(args.config)
    tvim.compile()


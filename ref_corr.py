"""References correction script."""
import re
import argparse


def ref_corr(text):

    def upper_surname(match):
        return match.group(1).upper() + r', ' + match.group(2).upper() + r'.'

    text = re.sub(r'(\d+)\s*-+\s*(\d+)', r'\1--\2', text, 0, re.IGNORECASE)
    # English
    text = re.sub(r'(P.)\s+(\d+)', r'P.\;\2', text, 0, re.IGNORECASE)
    text = re.sub(r'(\d+)\s+(p.?)', r'\1\;p.', text, 0, re.IGNORECASE)
    text = re.sub(r'([a-zA-Z]+)[,\\;\s]+([A-Z]{1}[a-z]?)\.{1}', upper_surname,
                  text)
    # Russian
    text = re.sub(r'(С.)\s+(\d+)', r'С.\;\2', text, 0, re.IGNORECASE)
    text = re.sub(r'(\d+)\s+(с.?)', r'\1\;с.', text, 0, re.IGNORECASE)
    return text


def corr(text):
    text = ref_corr(text)
    return text


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='TVIM reference corrector')

    input_path = parser.add_argument('--input', '-I', type=str,
                                     default='tmp/i_corr.txt',
                                     help='input file')
    output_path = parser.add_argument('--output', '-O', type=str,
                                      default='tmp/o_corr.txt',
                                      help='output file')
    _args = parser.parse_args()

    with open(_args.input, 'rt') as file:
        _text = file.read()

    _text = corr(_text)

    with open(_args.output, 'wt') as file:
        file.write(_text)

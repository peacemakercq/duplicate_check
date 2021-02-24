import os
import sys
import getopt
from win32com import client as wc
from functools import cmp_to_key
import re
import editdistance

SEPARATOR = "---------------------------------------------------------------------------------"

CATEGORY_PATTERN = re.compile(r'\b计算题\b|\b判断题\b|\b简单题\b|\b单选题\b|\b填空题\b')
STEM_PATTERN = re.compile(r'^([0-9]+)[\.． ：\s\t题]*(.*)[^0-9]$')
CHINESE_PATTERN = re.compile(r'[\u4e00-\u9fa5]')
DISTANCE_THRESHOLD = 1


def extract_chinese(raw):
    return "".join(CHINESE_PATTERN.findall(raw))


def get_word_files(rootDir):
    for dirName, subdirList, fielList in os.walk(rootDir):
        for fname in fielList:
            if fname.endswith('doc') or fname.endswith('docx'):
                yield os.path.abspath(os.path.join(dirName, fname))
        for dir in subdirList:
            get_word_files(dir)


def get_word_text(word, wordFile):
    tempFile = os.path.splitext(wordFile)[0] + '.txt'
    doc = word.Documents.Open(wordFile)
    doc.SaveAs(tempFile, 4, False, "", True, "", False, False, False, False)
    doc.Close()
    file = open(tempFile, mode='r', encoding="GB2312", errors='ignore')
    result = file.read()
    file.close()
    os.remove(tempFile)
    return result


def get_exam_papers(rootDir):
    word = wc.Dispatch('Word.Application')
    word.visible = False
    for fword in get_word_files(rootDir):
        text = get_word_text(word, fword)
        yield (fword, text)
    word.Quit()


def get_question(line):
    m = STEM_PATTERN.match(line)
    if not m:
        return None
    no, stem = m.groups()

    if not stem.strip():
        return None

    return (no, stem.strip())


def is_category(line):
    # return line.strip().endswith('题')
    return CATEGORY_PATTERN.match(line.strip())


def get_questions(text):
    category = '未分类'
    for line in text.splitlines():
        if not line.strip():
            continue
        question = get_question(line)
        if question:
            no, stem = question
            yield (category, no, stem, extract_chinese(stem))
        elif is_category(line):
            category = line.strip()


def debug_write_line(r, questions):
    for question in questions:
        category, no, stem, _ = question
        write_line(r, "[%s]" % category, "[%s]" % no, stem)


# def compare_questions(question1, question2):
#     category1, _, _, ch_stem1 = question1
#     category2, _, _, ch_stem2 = question2
#     if category1 < category2:
#         return -1
#     elif category1 > category2:
#         return 1
#     elif ch_stem1 < ch_stem2:
#         return -1
#     elif ch_stem1 > ch_stem2:
#         return 1
#     else:
#         return 0

# def sort_questions_by_category(questions):
#     return sorted(questions, key=cmp_to_key(compare_questions))

def is_similar(question1, question2):
    category1, _, _, ch_stem1 = question1
    category2, _, _, ch_stem2 = question2

    if category1 != category2:
        return False

    return editdistance.eval(ch_stem1, ch_stem2) < DISTANCE_THRESHOLD


def get_similar_questions(source, questions):
    for question in questions:
        if is_similar(source, question):
            yield question


def write_line(f, *words):
    for word in words:
        f.write(str(word))
    f.write('\n')


def main(argv):
    total_papers = 0
    total_questions = 0
    total_real_questions = 0

    inputdir = 'data_debug'
    outputdir = 'output'

    try:
        opts, _ = getopt.getopt(argv, "hi:o:", ["input=", "output="])
    except getopt.GetoptError:
        print('Usage: duplicate_check.py -i <input> -o <output>')
        sys.exit(2)

    for opt, arg in opts:
        if opt == '-h':
            print('Usage: duplicate_check.py -i <input> -o <output>')
            sys.exit()
        elif opt in ("-i", "--input"):
            inputdir = arg
        elif opt in ("-o", "--output"):
            outputdir = arg

    print('Input directory: ', inputdir)
    print('Output directoy: ', outputdir)

    if not os.path.exists(outputdir):
        os.mkdir(outputdir)

    summaryfile = os.path.join(outputdir, "汇总结果.txt")
    with open(summaryfile, 'w', encoding='utf-8', errors='ignore') as summary:
        for paper in get_exam_papers(inputdir):
            path, content = paper
            print('Processing...', path)

            # paper_name = os.path.splitext(os.path.basename(path))[0].strip()
            paper_name = os.path.splitext(os.path.relpath(path, inputdir))[0].strip().replace('\\', '_').replace('（','(').replace('）',')')

            outfile = os.path.join(outputdir, '[分析结果] ' + paper_name + '.txt')

            with open(outfile, 'w', encoding='utf-8', errors='ignore') as result:

                questions = list(get_questions(content))

                # sorted_questions = list(sort_questions_by_category(questions))
                # debug_write_line(r, sorted_questions)

                questions_unsearched = list(questions)
                questions_unsearched.reverse()
                similar_questions_collection = []
                while len(questions_unsearched) > 1:
                    question = questions_unsearched.pop()
                    similar_questions = list(get_similar_questions(
                        question, questions_unsearched))
                    if similar_questions:
                        for similar in similar_questions:
                            questions_unsearched.remove(similar)
                        similar_questions.append(question)
                        similar_questions.reverse()
                        similar_questions_collection.append(similar_questions)

                similar_questions_collection = sorted(
                    similar_questions_collection, key=lambda x: len(x), reverse=True)
                duplicate_count = 0
                for collection in similar_questions_collection:
                    duplicate_count += len(collection) - 1

                question_count = len(questions)
                real_question_count = len(questions) - duplicate_count
                duplicate_ratio = (question_count - real_question_count) / question_count

                write_line(result, SEPARATOR)
                write_line(result, "试卷名称：", paper_name)
                write_line(result, SEPARATOR)
                write_line(result, "题目总数：", question_count)
                write_line(result, "存在相似题数量: ", len(
                    similar_questions_collection))
                write_line(result, "去除相似题后数量: ", real_question_count)
                write_line(result, "重复率: ", "{0:.2%}".format(duplicate_ratio))

                write_line(result, SEPARATOR)
                write_line(result, "相似题（按重复次数由高到低排序）：")
                write_line(result, SEPARATOR)
                for collection in similar_questions_collection:
                    write_line(result, "重复", len(collection), "遍: ")
                    debug_write_line(result, collection)

                total_papers += 1
                total_questions += question_count
                total_real_questions += real_question_count

                write_line(summary, paper_name, ': ', real_question_count, '/',
                           question_count, '(', '{0:.2%}'.format(duplicate_ratio), ')')

            # break
        write_line(summary, SEPARATOR)
        write_line(summary, "处理试卷总数: ", total_papers)
        write_line(summary, "总题目数: ", total_questions)
        write_line(summary, "总去除相似题数: ", total_real_questions)
        write_line(summary, "总重复率: ", "{0:.2%}".format(
            (total_questions - total_real_questions) / total_questions))
        write_line(summary, SEPARATOR)


if __name__ == "__main__":
    main(sys.argv[1:])

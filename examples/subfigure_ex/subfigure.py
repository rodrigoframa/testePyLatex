from pylatex import Document, Section, Figure, SubFigure, NoEscape
import os

if __name__ == '__main__':
    doc = Document(default_filepath='subfigures')
    image_filename = os.path.join(os.path.dirname(__file__), 'logo.png')

    with doc.create(Section('Showing subfigures')):
        with doc.create(Figure(position='h!')) as kittens:
            with doc.create(SubFigure(
                    position='b',
                    width=NoEscape(r'0.45\linewidth'))) as left_kitten:

                left_kitten.add_image(image_filename,
                                      width=NoEscape(r'\linewidth'))
                left_kitten.add_caption('Kitten on the left')
            with doc.create(SubFigure(
                    position='b',
                    width=NoEscape(r'0.45\linewidth'))) as right_kitten:

                right_kitten.add_image(image_filename,
                                       width=NoEscape(r'\linewidth'))
                right_kitten.add_caption('Kitten on the right')
            kittens.add_caption("Two kittens")

    doc.generate_pdf(clean_tex=False)
"""
Sudoku
Shivneel Mistry
05/05/2021
"""
from random import shuffle, choice, randint
import sys
from typing import Union, List
import xlsxwriter
import pygame

WIDTH, HEIGHT = 900, 950
CHIN = 50
WHITE = (255, 255, 255)
GREY = (160, 160, 160)
BLACK = (0, 0, 0)
GRID = [["", "", "", "", "", "", "", "", ""] for _ in range(9)]
SOLVED = []
VALS = 0


def run_visualisation() -> None:
    """
    Display a Sudoku Board
    :return: None
    """
    # initialize pygame
    pygame.init()
    # create screen
    screen = pygame.display.set_mode((WIDTH, HEIGHT))
    # rename window title
    pygame.display.set_caption("Sudoku")
    create_board()

    # Print solved example, and current grid to console
    print("current unsolved GRID")
    for i in GRID:
        print(i)

    print()
    print("solved example of GRID")
    for i in SOLVED:
        print(i)
    # run even loop
    event_loop(screen)


def draw_board(screen: pygame.Surface) -> None:
    """
    Draw all required components of a Sudoku Board
    :param screen: pygame.Surface
    :return: None
    """
    # fill background
    screen.fill(WHITE)
    # draw outer rectangle
    pygame.draw.rect(screen, BLACK, pygame.Rect(0, 0, WIDTH, HEIGHT), 5)
    # draw all vertical lines and horizontal
    for i in range(1, 9):
        line_width = 5 if i % 3 == 0 else 1
        pygame.draw.line(screen, BLACK, ((WIDTH // 9) * i, 0),
                         ((WIDTH // 9) * i, (HEIGHT - CHIN)), line_width)
        pygame.draw.line(screen, BLACK, (0, ((HEIGHT - CHIN) // 9) * i),
                         (WIDTH, ((HEIGHT - CHIN) // 9) * i), line_width)

    pygame.draw.line(screen, BLACK, (0, ((HEIGHT - CHIN) // 9) * 9),
                     (WIDTH, ((HEIGHT - CHIN) // 9) * 9), 1)

    pygame.draw.rect(screen, GREY, pygame.Rect(WIDTH // 4, (((HEIGHT - CHIN) // 9) * 9), 150, 50), 2)
    pygame.draw.rect(screen, GREY, pygame.Rect(WIDTH // 2 + 75, (((HEIGHT - CHIN) // 9) * 9), 150, 50), 2)

    font = pygame.font.SysFont('Consolas', 15)
    save_text_img = font.render("save as image", True, pygame.Color('black'))
    save_text_xlsx = font.render("save as xlsx", True, pygame.Color('black'))

    screen.blit(save_text_img, (WIDTH // 4 + 20, (((HEIGHT - CHIN) // 9) * 9) + 20))
    screen.blit(save_text_xlsx, (WIDTH // 2 + 100, (((HEIGHT - CHIN) // 9) * 9) + 20))


def draw_numbers(screen: pygame.Surface) -> None:
    """
    Displays all numbers from Grid created onto the Puzzle board
    :param screen:
    :return: None
    """
    for i in range(9):
        for j in range(9):
            val = GRID[i][j]
            if val != '' and 0 < int(val) < 10:
                font = pygame.font.SysFont('Consolas', 40)
                text = font.render(str(val), True, pygame.Color('black'))
                screen.blit(text, (100 * j + 40, 100 * i + 40))


def create_board() -> None:
    """
    Creates A valid Random Board
    :return: None
    """
    temp = list(range(1, 10))
    shuffle(temp)
    count = 0
    for i in range(3):
        for j in range(3):
            GRID[i][j] = temp[count]
            count += 1

    # horizontal
    # obtain each row from first quadrant
    tr = GRID[0][:3]
    mr = GRID[1][:3]
    br = GRID[2][:3]
    # assigning the row to be the the values of the other rows (2 from the same 1 different)
    mid_r = [tr[0], tr[2], br[1]]
    top_r = [br[0], br[2], mr[1]]
    bot_r = [tr[1], mr[0], mr[2]]
    # shuffle for randomness
    shuffle(mid_r), shuffle(top_r), shuffle(bot_r)
    # assign it
    GRID[0][3:6] = top_r
    GRID[1][3:6] = mid_r
    GRID[2][3:6] = bot_r
    # find difference for last row
    last_t, last_m, last_b = list(set(temp).difference(set(GRID[0][0:6]))), list(
        set(temp).difference(set(GRID[1][0:6]))), list(set(temp).difference(set(GRID[2][0:6])))
    # randomize it
    shuffle(last_t), shuffle(last_m), shuffle(last_b)
    # assign it
    GRID[0][6:9] = last_t
    GRID[1][6:9] = last_m
    GRID[2][6:9] = last_b

    # vertical
    # obtain cols for first quadrant
    mc = [GRID[0][1], GRID[1][1], GRID[2][1]]
    lc = [GRID[0][0], GRID[1][0], GRID[2][0]]
    rc = [GRID[0][2], GRID[1][2], GRID[2][2]]
    # assigning the row to be the the values of the other rows (2 from the same 1 different)
    mid_c = [lc[0], lc[2], rc[1]]
    left_c = [rc[0], rc[2], mc[1]]
    right_c = [lc[1], mc[0], mc[2]]
    shuffle(mid_c), shuffle(left_c), shuffle(right_c)
    # assign grid to values
    count = 0
    for i in range(3, 6):
        GRID[i][0] = left_c[count]
        GRID[i][1] = mid_c[count]
        GRID[i][2] = right_c[count]
        count += 1
    # obtain last quadrant remaining values
    last_lc, last_mc, last_rc = list(set(temp).difference(set([GRID[i][0] for i in range(0, 6)]))), list(
        set(temp).difference(set([GRID[i][1] for i in range(0, 6)]))), list(set(temp).difference(set([GRID[i][2] for i
                                                                                                      in range(0, 6)])))
    # shuffle them
    shuffle(last_lc), shuffle(last_mc), shuffle(last_rc)
    # assign the last quadrant to these values
    count = 0
    for i in range(6, 9):
        GRID[i][0] = last_lc[count]
        GRID[i][1] = last_mc[count]
        GRID[i][2] = last_rc[count]
        count += 1

    # start backtrack solve for remaining 36 cells (4 quadrants)
    backtrack_solve(GRID)
    SOLVED.extend([[i for i in j] for j in GRID])
    # remove values for user to play
    unsolve_board()


def unsolve_board():
    """
    Remove filled in values to generate random Sudoku Board (may have multiple Solutions)
    :return: None
    """
    # we want 21-33 filled in blocks
    remove = randint(48, 60)
    indices = [[i, j] for j in range(0, 9) for i in range(0, 9)]
    for i in range(remove):
        pos = choice(indices)
        (GRID[pos[0]])[pos[1]] = ""
        indices.remove(pos)


def backtrack_solve(board: List[List[int]]) -> bool:
    """
    Return if puzzle is solvable with current value
    :param board: List[List[int]]
    :return: bool
    """
    i, j = next_cell(board)
    if i == "":
        return True
    temp = list(range(1, 10))
    shuffle(temp)
    for k in temp:
        if allowed(board, i, j, k):
            board[i][j] = k
            if backtrack_solve(board):
                return True
            board[i][j] = ""
    return False


def next_cell(board: List[List[int]]) -> Union[tuple[int, int], tuple[str, str]]:
    """
    Return tuple of x,y Coordinates of next empty cell.
    :param board: List[List[int]]
    :return: Tuple(int, int) or Tuple(str, str)
    """
    for x in range(0, 9):
        for y in range(0, 9):
            if board[x][y] == "":
                return x, y
    return "", ""


def allowed(board: List[List[int]], row: int, col: int, val: int) -> bool:
    """
    Returns True if val is unique in column, row and its grid
    :param board:
    :param row: int
    :param col: int
    :param val: int
    :return: bool
    """
    # check if val in row
    for i in range(9):
        if board[row][i] == val:
            return False

    # check if val in column
    for j in range(9):
        if board[j][col] == val:
            return False

    # check if val is in grid
    row_index = (row // 3) * 3
    col_index = (col // 3) * 3
    for i in range(row_index, row_index + 3):
        for j in range(col_index, col_index + 3):
            if GRID[i][j] == val:
                return False
    # default true
    return True


def save_as_xslx() -> None:
    """
    Saves file as XSLX file with solution
    :return: None
    """
    workbook = xlsxwriter.Workbook('SudokuGame.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write(10, 6, "Sudoku Puzzle")
    border_format = workbook.add_format({
        'border': 1,
        'align': 'left',
        'font_size': 10
    })
    worksheet.conditional_format('D13:L21', {'type': 'no_blanks', 'format': border_format})

    row = 0
    while row < 9:
        col = 0
        while col < 9:
            worksheet.write(row + 12, col + 3, (GRID[row][col]))
            col += 1
        row += 1

    worksheet.write(24, 6, "Sudoku Puzzle Solution")
    worksheet.conditional_format('D27:L35', {'type': 'no_blanks', 'format': border_format})
    sol_row = 0
    while sol_row < 9:
        sol_col = 0
        while sol_col < 9:
            worksheet.write(sol_row + 26, sol_col + 3, (SOLVED[sol_row][sol_col]))
            sol_col += 1
        sol_row += 1

    workbook.close()


def event_loop(screen: pygame.Surface) -> None:
    """
    Responds to events and updates the display.
    :param screen: pygame.Surface
    :return: None
    """
    while True:
        mouse = pygame.mouse.get_pos()
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                sys.exit()

            if event.type == pygame.MOUSEBUTTONDOWN:
                if WIDTH // 4 <= mouse[0] <= (WIDTH // 4) + 150 and (((HEIGHT - CHIN) // 9) * 9) <= mouse[1] <= (
                        ((HEIGHT - CHIN) // 9) * 9) + 50:
                    pygame.image.save(screen, "Sudoku Puzzle.jpeg")
                    print("Sudoku Puzzle.jpeg saved to project folder")

                if WIDTH // 2 + 75 <= mouse[0] <= (WIDTH // 2 + 75) + 150 and (((HEIGHT - CHIN) // 9) * 9) <= mouse[
                    1] <= (((HEIGHT - CHIN) // 9) * 9) + 50:
                    save_as_xslx()
                    print("SudokuGame.xlsx saved to project folder")


        draw_board(screen)
        draw_numbers(screen)
        pygame.display.flip()


if __name__ == '__main__':
    run_visualisation()

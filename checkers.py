import pygame
import numpy as np

# Initialize Pygame
pygame.init()

# Constants
WINDOW_SIZE = 800
BOARD_SIZE = 8
SQUARE_SIZE = WINDOW_SIZE // BOARD_SIZE
BLACK = (0, 0, 0)
WHITE = (255, 255, 255)
RED = (255, 0, 0)
GRAY = (128, 128, 128)

# Set up the display
screen = pygame.display.set_mode((WINDOW_SIZE, WINDOW_SIZE))
pygame.display.set_caption('Checkers')

class Board:
    def __init__(self):
        self.board = np.zeros((BOARD_SIZE, BOARD_SIZE))
        self.selected_piece = None
        self.setup_board()

    def setup_board(self):
        # 1 represents black pieces, 2 represents red pieces
        for row in range(BOARD_SIZE):
            for col in range(BOARD_SIZE):
                if row < 3 and (row + col) % 2 == 1:
                    self.board[row][col] = 1
                elif row > 4 and (row + col) % 2 == 1:
                    self.board[row][col] = 2

    def draw(self, screen):
        for row in range(BOARD_SIZE):
            for col in range(BOARD_SIZE):
                # Draw board squares
                color = WHITE if (row + col) % 2 == 0 else BLACK
                pygame.draw.rect(screen, color, 
                               (col * SQUARE_SIZE, row * SQUARE_SIZE, 
                                SQUARE_SIZE, SQUARE_SIZE))
                
                # Draw pieces
                if self.board[row][col] != 0:
                    color = BLACK if self.board[row][col] == 1 else RED
                    center = (col * SQUARE_SIZE + SQUARE_SIZE // 2,
                            row * SQUARE_SIZE + SQUARE_SIZE // 2)
                    pygame.draw.circle(screen, color, center, SQUARE_SIZE // 2 - 10)

    def get_piece_at_pos(self, pos):
        col = pos[0] // SQUARE_SIZE
        row = pos[1] // SQUARE_SIZE
        if 0 <= row < BOARD_SIZE and 0 <= col < BOARD_SIZE:
            return (row, col)
        return None

    def is_valid_move(self, start, end):
        if not (0 <= end[0] < BOARD_SIZE and 0 <= end[1] < BOARD_SIZE):
            return False
        
        if self.board[end[0]][end[1]] != 0:
            return False

        piece = self.board[start[0]][start[1]]
        row_diff = end[0] - start[0]
        col_diff = abs(end[1] - start[1])

        # Basic movement rules
        if piece == 1 and row_diff <= 0:  # Black pieces move down
            return False
        if piece == 2 and row_diff >= 0:  # Red pieces move up
            return False
        
        if abs(row_diff) == 1 and col_diff == 1:
            return True
        
        return False

    def move_piece(self, start, end):
        if self.is_valid_move(start, end):
            self.board[end[0]][end[1]] = self.board[start[0]][start[1]]
            self.board[start[0]][start[1]] = 0
            return True
        return False

def main():
    board = Board()
    running = True
    selected_piece = None

    while running:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                running = False
            
            if event.type == pygame.MOUSEBUTTONDOWN:
                pos = pygame.mouse.get_pos()
                clicked_pos = board.get_piece_at_pos(pos)
                
                if clicked_pos is not None:
                    if selected_piece is None:
                        if board.board[clicked_pos[0]][clicked_pos[1]] != 0:
                            selected_piece = clicked_pos
                    else:
                        if board.move_piece(selected_piece, clicked_pos):
                            selected_piece = None
                        else:
                            selected_piece = None

        # Draw the game
        screen.fill(WHITE)
        board.draw(screen)
        
        # Highlight selected piece
        if selected_piece:
            pygame.draw.rect(screen, GRAY,
                           (selected_piece[1] * SQUARE_SIZE,
                            selected_piece[0] * SQUARE_SIZE,
                            SQUARE_SIZE, SQUARE_SIZE), 4)
        
        pygame.display.flip()

    pygame.quit()

if __name__ == "__main__":
    main()
